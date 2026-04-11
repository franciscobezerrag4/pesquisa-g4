"""
Dashboard interativo da Pesquisa G4 — integrado ao admin do sistema.

Lê direto do SQLite (pesquisa.db) e renderiza HTML com filtros dinâmicos.
Montado em app.py via rota /admin/dashboard (protegida por auth).
"""

from collections import defaultdict
from datetime import datetime


def computar_nps(notas):
    total = len(notas)
    if total == 0:
        return {'total': 0, 'media': 0.0, 'promotores': 0, 'neutros': 0,
                'detratores': 0, 'pct_prom': 0.0, 'pct_det': 0.0, 'nps': 0.0}
    promotores = sum(1 for n in notas if n >= 9)
    neutros = sum(1 for n in notas if 7 <= n <= 8)
    detratores = sum(1 for n in notas if n <= 6)
    pct_p = promotores / total * 100
    pct_d = detratores / total * 100
    return {
        'total': total,
        'media': sum(notas) / total,
        'promotores': promotores,
        'neutros': neutros,
        'detratores': detratores,
        'pct_prom': pct_p,
        'pct_det': pct_d,
        'nps': pct_p - pct_d,
    }


def classificar_zona(nps):
    if nps >= 50:
        return 'excelencia'
    if nps >= 0:
        return 'atencao'
    return 'critica'


def build_dashboard_data(conn, senioridade='ambos', excluir_propria=False, anos_selecionados=None):
    """Lê dados do SQLite e monta o dicionário para o template."""
    if anos_selecionados is None:
        anos_selecionados = set()

    # Fetch raw data
    nps_rows = conn.execute('''
        SELECT r.id, r.area_respondente, r.senioridade, r.ano_entrada,
               n.area_avaliada, n.nota, n.comentario
        FROM respostas r
        JOIN nps_areas n ON n.resposta_id = r.id
    ''').fetchall()

    ind_rows = conn.execute('''
        SELECT r.id, r.area_respondente, r.senioridade, r.ano_entrada,
               a.funcionario_avaliado, a.recomenda, a.motivo
        FROM respostas r
        JOIN avaliacoes_individuais a ON a.resposta_id = r.id
    ''').fetchall()

    total_enviados = conn.execute('SELECT COUNT(*) FROM tokens').fetchone()[0]

    # Anos disponíveis
    anos_todos_set = set()
    for r in nps_rows:
        if r['ano_entrada']:
            anos_todos_set.add(str(r['ano_entrada']))
    anos_disponiveis = sorted(anos_todos_set)

    def match_senior(senior_val):
        if senioridade == 'ambos':
            return True
        if senioridade == 'lideranca':
            return senior_val == 'Liderança'
        if senioridade == 'ic':
            return senior_val == 'Contribuidor Individual'
        return True

    def match_ano(ano_val):
        if not anos_selecionados:
            return True
        return str(ano_val) in anos_selecionados

    # Filtrar NPS rows
    nps_filt = []
    for r in nps_rows:
        if not match_senior(r['senioridade']):
            continue
        if not match_ano(r['ano_entrada']):
            continue
        if excluir_propria and r['area_respondente'] == r['area_avaliada']:
            continue
        nps_filt.append(r)

    # Filtrar ind rows
    ind_filt = [r for r in ind_rows if match_senior(r['senioridade']) and match_ano(r['ano_entrada'])]

    total_respostas = len(set(r['id'] for r in nps_filt))

    # NPS por Area
    por_area = defaultdict(list)
    for r in nps_filt:
        if r['area_avaliada'] and r['nota'] is not None:
            por_area[r['area_avaliada']].append(int(r['nota']))

    areas = []
    for nome, notas in por_area.items():
        stats = computar_nps(notas)
        areas.append({'nome': nome, **stats, 'zona': classificar_zona(stats['nps'])})
    areas_sorted = sorted(areas, key=lambda a: a['nps'], reverse=True)

    total_av = sum(a['total'] for a in areas)
    nps_geral = sum(a['nps'] * a['total'] for a in areas) / total_av if total_av else 0
    media_geral = sum(a['media'] * a['total'] for a in areas) / total_av if total_av else 0

    zonas_count = defaultdict(int)
    for a in areas:
        zonas_count[a['zona']] += 1

    # Breakdown: votos por área de origem
    votos_origem = defaultdict(lambda: defaultdict(list))
    for r in nps_filt:
        if r['area_avaliada'] and r['nota'] is not None and r['area_respondente']:
            votos_origem[r['area_avaliada']][r['area_respondente']].append(int(r['nota']))

    votos_por_area_origem = {}
    for area_aval, por_origem in votos_origem.items():
        itens = []
        for origem, notas in por_origem.items():
            stats = computar_nps(notas)
            itens.append({
                'origem': origem,
                'total': stats['total'],
                'nps': stats['nps'],
                'media': stats['media'],
                'promotores': stats['promotores'],
                'detratores': stats['detratores'],
            })
        votos_por_area_origem[area_aval] = sorted(itens, key=lambda x: -x['total'])

    # Divergência Liderança x IC (sempre ambos os grupos, respeita excluir_propria + ano)
    por_area_senior = defaultdict(lambda: {'Liderança': [], 'Contribuidor Individual': []})
    for r in nps_rows:
        if excluir_propria and r['area_respondente'] == r['area_avaliada']:
            continue
        if not match_ano(r['ano_entrada']):
            continue
        senior = r['senioridade']
        if senior in ('Liderança', 'Contribuidor Individual') and r['area_avaliada'] and r['nota'] is not None:
            por_area_senior[r['area_avaliada']][senior].append(int(r['nota']))

    divergencias = []
    for area, grupos in por_area_senior.items():
        lid_notas = grupos['Liderança']
        ic_notas = grupos['Contribuidor Individual']
        if lid_notas and ic_notas:
            lid_s = computar_nps(lid_notas)
            ic_s = computar_nps(ic_notas)
            divergencias.append({
                'area': area,
                'nps_lideranca': lid_s['nps'], 'n_lid': lid_s['total'],
                'nps_ic': ic_s['nps'], 'n_ic': ic_s['total'],
                'gap': lid_s['nps'] - ic_s['nps'],
            })
    divergencias_sorted = sorted(divergencias, key=lambda d: abs(d['gap']), reverse=True)

    # Pessoas
    pessoas_agg = defaultdict(lambda: {'rec': 0, 'nao_rec': 0})
    for r in ind_filt:
        func = r['funcionario_avaliado']
        if not func:
            continue
        if r['recomenda'] == 1:
            pessoas_agg[func]['rec'] += 1
        elif r['recomenda'] == 0:
            pessoas_agg[func]['nao_rec'] += 1

    pessoas = []
    for nome, c in pessoas_agg.items():
        total = c['rec'] + c['nao_rec']
        pct = c['rec'] / total * 100 if total else 0
        pessoas.append({'nome': nome, 'total': total, 'rec': c['rec'],
                        'nao_rec': c['nao_rec'], 'pct': pct})

    pessoas_ranking = sorted([p for p in pessoas if p['total'] >= 2],
                             key=lambda p: (-p['pct'], -p['total']))
    alertas = sorted([p for p in pessoas if p['pct'] < 50 and p['total'] >= 2],
                     key=lambda p: p['pct'])

    # Comentários por area
    coment_por_area = defaultdict(list)
    for r in nps_filt:
        coment = r['comentario']
        if coment and str(coment).strip() and r['area_avaliada']:
            coment_por_area[r['area_avaliada']].append({
                'area_respondente': r['area_respondente'] or '',
                'senioridade': r['senioridade'] or '',
                'texto': str(coment).strip(),
            })

    # Motivos não-rec por pessoa
    motivos_nao_rec = defaultdict(list)
    for r in ind_filt:
        motivo = r['motivo']
        if r['recomenda'] == 0 and motivo and str(motivo).strip():
            motivos_nao_rec[r['funcionario_avaliado']].append({
                'area_respondente': r['area_respondente'] or '',
                'senioridade': r['senioridade'] or '',
                'motivo': str(motivo).strip(),
            })

    # URLs (preserva estado ao trocar filtros)
    exc = '1' if excluir_propria else '0'
    toggle_exc = '0' if excluir_propria else '1'
    anos_str = ','.join(sorted(anos_selecionados)) if anos_selecionados else ''

    base_url = '/admin/dashboard'
    urls = {
        'senior_ambos': f'{base_url}?senioridade=ambos&excluir_propria={exc}&anos={anos_str}',
        'senior_lider': f'{base_url}?senioridade=lideranca&excluir_propria={exc}&anos={anos_str}',
        'senior_ic': f'{base_url}?senioridade=ic&excluir_propria={exc}&anos={anos_str}',
        'toggle_excluir': f'{base_url}?senioridade={senioridade}&excluir_propria={toggle_exc}&anos={anos_str}',
        'anos_todos': f'{base_url}?senioridade={senioridade}&excluir_propria={exc}&anos=',
        'anos_toggle': {},
        'reset': base_url,
    }

    for y in anos_disponiveis:
        novo_set = set(anos_selecionados)
        if y in novo_set:
            novo_set.remove(y)
        else:
            novo_set.add(y)
        novo_str = ','.join(sorted(novo_set)) if novo_set else ''
        urls['anos_toggle'][y] = f'{base_url}?senioridade={senioridade}&excluir_propria={exc}&anos={novo_str}'

    return {
        'modificado': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'total_respostas': total_respostas,
        'total_enviados': total_enviados,
        'taxa': (total_respostas / total_enviados * 100) if total_enviados else 0,
        'nps_geral': nps_geral,
        'media_geral': media_geral,
        'areas': areas_sorted,
        'zonas_count': dict(zonas_count),
        'divergencias': divergencias_sorted,
        'pessoas_ranking': pessoas_ranking,
        'alertas': alertas,
        'total_pessoas': len(pessoas),
        'coment_por_area': dict(coment_por_area),
        'motivos_nao_rec': dict(motivos_nao_rec),
        'votos_por_area_origem': votos_por_area_origem,
        'anos_disponiveis': anos_disponiveis,
        'filtros': {
            'senioridade': senioridade,
            'excluir_propria': excluir_propria,
            'anos': sorted(anos_selecionados),
        },
        'urls': urls,
    }


DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard — Pesquisa G4</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
  :root {
    --cor-excelencia: #198754;
    --cor-atencao: #fd7e14;
    --cor-critica: #dc3545;
  }
  body { background: #f5f6f8; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  .navbar-brand { font-weight: 700; letter-spacing: -0.3px; }
  .kpi-card { border: none; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); transition: transform 0.15s; }
  .kpi-card:hover { transform: translateY(-2px); }
  .kpi-value { font-size: 2rem; font-weight: 700; line-height: 1; }
  .kpi-label { font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.5px; color: #6c757d; }
  .zona-excelencia { color: var(--cor-excelencia); }
  .zona-atencao { color: var(--cor-atencao); }
  .zona-critica { color: var(--cor-critica); }
  .badge-excelencia { background: var(--cor-excelencia); }
  .badge-atencao { background: var(--cor-atencao); }
  .badge-critica { background: var(--cor-critica); }
  .nav-tabs .nav-link { color: #495057; font-weight: 500; border: none; border-bottom: 3px solid transparent; }
  .nav-tabs .nav-link.active { color: #0d6efd; border-bottom-color: #0d6efd; background: transparent; }
  .area-card { cursor: pointer; border: 1px solid #e9ecef; border-radius: 10px; padding: 16px; margin-bottom: 12px; background: white; transition: all 0.2s; }
  .area-card:hover { border-color: #0d6efd; box-shadow: 0 2px 8px rgba(13,110,253,0.1); }
  .area-card.expanded { border-color: #0d6efd; background: #f8f9ff; }
  .area-nome { font-weight: 600; font-size: 1.05rem; }
  .nps-pill { display: inline-block; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 0.9rem; color: white; min-width: 60px; text-align: center; }
  .comentario { padding: 10px 14px; background: #f8f9fa; border-left: 3px solid #0d6efd; margin-bottom: 8px; border-radius: 4px; font-size: 0.92rem; }
  .comentario-meta { color: #6c757d; font-size: 0.78rem; margin-top: 4px; }
  .table-clean { background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
  .search-box { max-width: 340px; }
  .filter-btn { border-radius: 20px; margin-right: 6px; margin-bottom: 6px; }
  .pessoa-row { transition: background 0.1s; }
  .pessoa-row:hover { background: #f0f4ff; }
  .hidden { display: none !important; }
  .detail-toggle { float: right; font-size: 0.85rem; color: #0d6efd; }
  .action-card { background: white; border-radius: 10px; padding: 20px; margin-bottom: 16px; border-left: 4px solid; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
  .action-critica { border-left-color: var(--cor-critica); }
  .action-atencao { border-left-color: var(--cor-atencao); }
  .action-excelencia { border-left-color: var(--cor-excelencia); }
  .footer-info { font-size: 0.82rem; color: #6c757d; }
  .ranking-row { cursor: pointer; }
  .ranking-row:hover { background: #f0f4ff; }
  .ranking-row.active { background: #e7f1ff; }
  .breakdown-row td { background: #f8f9ff; padding: 16px 24px !important; border-top: 2px solid #0d6efd; }
  .breakdown-title { font-weight: 600; color: #0d6efd; margin-bottom: 10px; font-size: 0.95rem; }
  .breakdown-table { background: white; border-radius: 6px; overflow: hidden; font-size: 0.88rem; }
  .breakdown-table th { background: #e9ecef; font-weight: 600; padding: 6px 10px; font-size: 0.78rem; text-transform: uppercase; color: #6c757d; }
  .breakdown-table td { padding: 6px 10px; }
  .expand-icon { display: inline-block; width: 16px; color: #6c757d; transition: transform 0.2s; }
  .ranking-row.active .expand-icon { transform: rotate(90deg); color: #0d6efd; }
  h4 { font-weight: 600; color: #212529; }
  .section-title { font-size: 1.1rem; font-weight: 600; margin-top: 24px; margin-bottom: 12px; color: #495057; }
  .filter-bar { background: white; border-radius: 10px; padding: 12px 16px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); display: flex; gap: 16px; align-items: center; flex-wrap: wrap; }
  .filter-bar .filter-label { font-size: 0.82rem; font-weight: 600; color: #6c757d; text-transform: uppercase; letter-spacing: 0.5px; }
  .filter-bar .btn { font-size: 0.85rem; }
  .filter-active-badge { display: inline-block; background: #cff4fc; color: #055160; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 600; }
</style>
</head>
<body>

<nav class="navbar navbar-dark bg-dark">
  <div class="container-fluid">
    <span class="navbar-brand">📊 Pesquisa G4 — Dashboard (ao vivo)</span>
    <span class="text-light footer-info">
      Atualizado {{ data.modificado }}
      &nbsp;<a href="{{ data.urls.reset }}" class="btn btn-sm btn-outline-light ms-2">↻ Recarregar</a>
      &nbsp;<a href="/admin" class="btn btn-sm btn-outline-light">← Admin</a>
    </span>
  </div>
</nav>

<div class="container-fluid py-4 px-4">

  <!-- FILTROS -->
  <div class="filter-bar">
    <div>
      <span class="filter-label">Senioridade:</span>
      <div class="btn-group ms-2" role="group">
        <a href="{{ data.urls.senior_ambos }}" class="btn btn-sm {% if data.filtros.senioridade == 'ambos' %}btn-primary{% else %}btn-outline-primary{% endif %}">Todos</a>
        <a href="{{ data.urls.senior_lider }}" class="btn btn-sm {% if data.filtros.senioridade == 'lideranca' %}btn-primary{% else %}btn-outline-primary{% endif %}">Liderança</a>
        <a href="{{ data.urls.senior_ic }}" class="btn btn-sm {% if data.filtros.senioridade == 'ic' %}btn-primary{% else %}btn-outline-primary{% endif %}">Contribuidor Individual</a>
      </div>
    </div>
    <div>
      <span class="filter-label">Ano de entrada <small class="text-muted">(multi)</small>:</span>
      <div class="btn-group ms-2" role="group">
        <a href="{{ data.urls.anos_todos }}" class="btn btn-sm {% if not data.filtros.anos %}btn-primary{% else %}btn-outline-primary{% endif %}">Todos</a>
        {% for y in data.anos_disponiveis %}
        <a href="{{ data.urls.anos_toggle[y] }}" class="btn btn-sm {% if y in data.filtros.anos %}btn-primary{% else %}btn-outline-primary{% endif %}">{{ y }}</a>
        {% endfor %}
      </div>
    </div>
    <div>
      <span class="filter-label">Auto-avaliação de área:</span>
      <a href="{{ data.urls.toggle_excluir }}" class="btn btn-sm ms-2 {% if data.filtros.excluir_propria %}btn-warning{% else %}btn-outline-warning{% endif %}">
        {% if data.filtros.excluir_propria %}✓ Excluída{% else %}Incluída{% endif %}
      </a>
    </div>
    {% if data.filtros.senioridade != 'ambos' or data.filtros.excluir_propria or data.filtros.anos %}
    <div class="ms-auto">
      <span class="filter-active-badge">🔍 Filtro ativo</span>
    </div>
    {% endif %}
  </div>

  <!-- KPIs -->
  <div class="row g-3 mb-4">
    <div class="col-md-2 col-sm-4">
      <div class="kpi-card card p-3">
        <div class="kpi-label">NPS Geral</div>
        <div class="kpi-value {% if data.nps_geral >= 50 %}zona-excelencia{% elif data.nps_geral >= 0 %}zona-atencao{% else %}zona-critica{% endif %}">{{ "%.1f"|format(data.nps_geral) }}</div>
      </div>
    </div>
    <div class="col-md-2 col-sm-4">
      <div class="kpi-card card p-3">
        <div class="kpi-label">Respostas</div>
        <div class="kpi-value">{{ data.total_respostas }}<small class="text-muted" style="font-size: 1rem;">/{{ data.total_enviados }}</small></div>
        <small class="text-muted">{{ "%.1f"|format(data.taxa) }}% taxa</small>
      </div>
    </div>
    <div class="col-md-2 col-sm-4">
      <div class="kpi-card card p-3">
        <div class="kpi-label">Média (0-10)</div>
        <div class="kpi-value">{{ "%.2f"|format(data.media_geral) }}</div>
      </div>
    </div>
    <div class="col-md-2 col-sm-4">
      <div class="kpi-card card p-3">
        <div class="kpi-label">🟢 Excelência</div>
        <div class="kpi-value zona-excelencia">{{ data.zonas_count.get('excelencia', 0) }}</div>
        <small class="text-muted">áreas NPS ≥ 50</small>
      </div>
    </div>
    <div class="col-md-2 col-sm-4">
      <div class="kpi-card card p-3">
        <div class="kpi-label">🟡 Atenção</div>
        <div class="kpi-value zona-atencao">{{ data.zonas_count.get('atencao', 0) }}</div>
        <small class="text-muted">áreas NPS 0-50</small>
      </div>
    </div>
    <div class="col-md-2 col-sm-4">
      <div class="kpi-card card p-3">
        <div class="kpi-label">🔴 Crítica</div>
        <div class="kpi-value zona-critica">{{ data.zonas_count.get('critica', 0) }}</div>
        <small class="text-muted">áreas NPS &lt; 0</small>
      </div>
    </div>
  </div>

  <!-- Tabs -->
  <ul class="nav nav-tabs mb-3" id="tabs">
    <li class="nav-item"><a class="nav-link active" data-tab="ranking" href="#">Ranking</a></li>
    <li class="nav-item"><a class="nav-link" data-tab="areas" href="#">Áreas & Comentários</a></li>
    <li class="nav-item"><a class="nav-link" data-tab="divergencia" href="#">Divergência Líder vs IC</a></li>
    <li class="nav-item"><a class="nav-link" data-tab="pessoas" href="#">Pessoas</a></li>
    <li class="nav-item"><a class="nav-link" data-tab="acoes" href="#">Ações Sugeridas</a></li>
  </ul>

  <!-- RANKING -->
  <div class="tab-content" id="tab-ranking">
    <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
      <div>
        <button class="btn btn-sm btn-outline-secondary filter-btn" data-filter="todos">Todas</button>
        <button class="btn btn-sm btn-outline-success filter-btn" data-filter="excelencia">🟢 Excelência</button>
        <button class="btn btn-sm btn-outline-warning filter-btn" data-filter="atencao">🟡 Atenção</button>
        <button class="btn btn-sm btn-outline-danger filter-btn" data-filter="critica">🔴 Crítica</button>
      </div>
      <input type="text" class="form-control form-control-sm search-box" id="search-ranking" placeholder="🔍 Buscar área...">
    </div>
    <small class="text-muted d-block mb-2">💡 Clique em qualquer linha pra ver de quais áreas vieram os votos.</small>
    <div class="table-responsive table-clean">
      <table class="table table-hover mb-0">
        <thead class="table-light">
          <tr>
            <th style="width: 30px;"></th>
            <th>#</th><th>Área</th><th>Zona</th>
            <th class="text-end">NPS</th><th class="text-end">Média</th>
            <th class="text-end">Avaliações</th>
            <th class="text-end">% Promotores</th><th class="text-end">% Detratores</th>
          </tr>
        </thead>
        <tbody id="ranking-tbody">
          {% for a in data.areas %}
          <tr class="ranking-row" data-zona="{{ a.zona }}" data-nome="{{ a.nome|lower }}" data-idx="{{ loop.index0 }}" onclick="toggleBreakdown({{ loop.index0 }})">
            <td class="text-center"><span class="expand-icon">▶</span></td>
            <td class="text-muted">{{ loop.index }}</td>
            <td><strong>{{ a.nome }}</strong></td>
            <td><span class="badge badge-{{ a.zona }}">{% if a.zona == 'excelencia' %}🟢{% elif a.zona == 'atencao' %}🟡{% else %}🔴{% endif %}</span></td>
            <td class="text-end"><span class="nps-pill badge-{{ a.zona }}">{{ "%.1f"|format(a.nps) }}</span></td>
            <td class="text-end">{{ "%.2f"|format(a.media) }}</td>
            <td class="text-end">{{ a.total }}</td>
            <td class="text-end">{{ "%.1f"|format(a.pct_prom) }}%</td>
            <td class="text-end">{{ "%.1f"|format(a.pct_det) }}%</td>
          </tr>
          <tr class="breakdown-row hidden" id="breakdown-{{ loop.index0 }}" data-zona="{{ a.zona }}" data-nome="{{ a.nome|lower }}">
            <td colspan="9">
              <div class="breakdown-title">Votos para <strong>{{ a.nome }}</strong> vieram de:</div>
              {% set breakdown = data.votos_por_area_origem.get(a.nome, []) %}
              {% if breakdown %}
              <table class="breakdown-table table table-sm mb-0">
                <thead>
                  <tr>
                    <th>Área de Origem</th>
                    <th class="text-end">Votos</th>
                    <th class="text-end">NPS</th>
                    <th class="text-end">Média</th>
                    <th class="text-end">Promotores</th>
                    <th class="text-end">Detratores</th>
                  </tr>
                </thead>
                <tbody>
                  {% for b in breakdown %}
                  <tr>
                    <td>{{ b.origem }}{% if b.origem == a.nome %} <span class="badge bg-info">própria área</span>{% endif %}</td>
                    <td class="text-end">{{ b.total }}</td>
                    <td class="text-end"><strong class="{% if b.nps >= 50 %}zona-excelencia{% elif b.nps >= 0 %}zona-atencao{% else %}zona-critica{% endif %}">{{ "%.1f"|format(b.nps) }}</strong></td>
                    <td class="text-end">{{ "%.2f"|format(b.media) }}</td>
                    <td class="text-end">{{ b.promotores }}</td>
                    <td class="text-end">{{ b.detratores }}</td>
                  </tr>
                  {% endfor %}
                </tbody>
              </table>
              {% else %}
              <div class="text-muted fst-italic small">Sem votos registrados.</div>
              {% endif %}
            </td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

  <!-- ÁREAS -->
  <div class="tab-content hidden" id="tab-areas">
    <div class="mb-3">
      <input type="text" class="form-control form-control-sm search-box" id="search-areas" placeholder="🔍 Buscar área...">
      <small class="text-muted">Clica em qualquer card para expandir e ver os comentários dos respondentes.</small>
    </div>
    <div id="areas-container">
      {% for a in data.areas %}
      <div class="area-card" data-nome="{{ a.nome|lower }}" data-zona="{{ a.zona }}" onclick="toggleArea(this)">
        <div class="d-flex justify-content-between align-items-center">
          <div>
            <div class="area-nome">
              {% if a.zona == 'excelencia' %}🟢{% elif a.zona == 'atencao' %}🟡{% else %}🔴{% endif %}
              {{ a.nome }}
            </div>
            <div class="text-muted small mt-1">
              {{ a.total }} avaliações · Média {{ "%.2f"|format(a.media) }} · {{ a.promotores }} promotores / {{ a.neutros }} neutros / {{ a.detratores }} detratores
            </div>
          </div>
          <div class="text-end">
            <span class="nps-pill badge-{{ a.zona }}">NPS {{ "%.1f"|format(a.nps) }}</span>
            <div class="detail-toggle mt-1">▼ {{ data.coment_por_area.get(a.nome, [])|length }} comentários</div>
          </div>
        </div>
        <div class="area-details mt-3 hidden">
          {% set coments = data.coment_por_area.get(a.nome, []) %}
          {% if coments %}
            {% for c in coments %}
            <div class="comentario">
              "{{ c.texto }}"
              <div class="comentario-meta">— {{ c.senioridade }} de <strong>{{ c.area_respondente }}</strong></div>
            </div>
            {% endfor %}
          {% else %}
            <div class="text-muted small fst-italic">Nenhum comentário escrito até agora.</div>
          {% endif %}
        </div>
      </div>
      {% endfor %}
    </div>
  </div>

  <!-- DIVERGÊNCIA -->
  <div class="tab-content hidden" id="tab-divergencia">
    <div class="alert alert-info small">
      <strong>Como ler:</strong> Gap positivo = Liderança vê a área melhor que os Contribuidores Individuais. Gap negativo = IC vê melhor que Liderança. Quanto maior o <code>|gap|</code>, maior o desalinhamento de percepção — esses são os pontos de calibração mais urgentes.
      {% if data.filtros.senioridade != 'ambos' %}
      <br><strong>Nota:</strong> esta aba sempre considera <em>ambos</em> os grupos (Liderança e IC) — o filtro de senioridade não se aplica aqui.
      {% endif %}
    </div>
    <div class="table-responsive table-clean">
      <table class="table table-hover mb-0">
        <thead class="table-light">
          <tr>
            <th>Área</th>
            <th class="text-end">NPS Liderança</th>
            <th class="text-end">NPS IC</th>
            <th class="text-end">Gap</th>
            <th>Interpretação</th>
          </tr>
        </thead>
        <tbody>
          {% for d in data.divergencias %}
          <tr>
            <td><strong>{{ d.area }}</strong></td>
            <td class="text-end">{{ "%.1f"|format(d.nps_lideranca) }} <small class="text-muted">(n={{ d.n_lid }})</small></td>
            <td class="text-end">{{ "%.1f"|format(d.nps_ic) }} <small class="text-muted">(n={{ d.n_ic }})</small></td>
            <td class="text-end">
              <strong class="{% if d.gap|abs > 15 %}text-danger{% elif d.gap|abs > 5 %}text-warning{% else %}text-success{% endif %}">
                {{ "%+.1f"|format(d.gap) }}
              </strong>
            </td>
            <td>
              {% if d.gap > 15 %}⚠️ Liderança otimista demais
              {% elif d.gap < -15 %}⚠️ IC vê melhor que Liderança
              {% elif d.gap|abs > 5 %}Divergência moderada
              {% else %}✓ Alinhado{% endif %}
            </td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

  <!-- PESSOAS -->
  <div class="tab-content hidden" id="tab-pessoas">
    <div class="mb-3">
      <input type="text" class="form-control form-control-sm search-box" id="search-pessoas" placeholder="🔍 Buscar pessoa...">
      <small class="text-muted d-block mt-1">Ranking completo ({{ data.pessoas_ranking|length }} pessoas com pelo menos 2 avaliações). Total avaliadas: {{ data.total_pessoas }}.</small>
    </div>

    {% if data.alertas %}
    <div class="alert alert-warning">
      <strong>⚠️ {{ data.alertas|length }} pessoa(s) em alerta</strong> (< 50% de recomendação) — ver lista abaixo.
    </div>
    <div class="section-title">Alertas (&lt; 50% recomendação)</div>
    <div class="table-responsive table-clean mb-4">
      <table class="table table-hover mb-0">
        <thead class="table-light">
          <tr>
            <th>Pessoa</th><th class="text-end">% Recom.</th>
            <th class="text-end">Rec.</th><th class="text-end">Não Rec.</th><th class="text-end">Total</th>
            <th>Motivos</th>
          </tr>
        </thead>
        <tbody>
          {% for p in data.alertas %}
          <tr class="pessoa-row">
            <td><strong>{{ p.nome }}</strong></td>
            <td class="text-end"><span class="badge bg-danger">{{ "%.0f"|format(p.pct) }}%</span></td>
            <td class="text-end">{{ p.rec }}</td>
            <td class="text-end">{{ p.nao_rec }}</td>
            <td class="text-end">{{ p.total }}</td>
            <td>
              {% set motivos = data.motivos_nao_rec.get(p.nome, []) %}
              {% if motivos %}
                {% for m in motivos %}
                  <div class="small text-muted">"{{ m.motivo }}" — <em>{{ m.senioridade }} de {{ m.area_respondente }}</em></div>
                {% endfor %}
              {% else %}
                <span class="text-muted small fst-italic">sem motivo escrito</span>
              {% endif %}
            </td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
    {% endif %}

    <div class="section-title">Ranking completo</div>
    <div class="table-responsive table-clean">
      <table class="table table-hover mb-0">
        <thead class="table-light">
          <tr>
            <th>#</th><th>Pessoa</th>
            <th class="text-end">% Recomendação</th>
            <th class="text-end">Recomenda</th>
            <th class="text-end">Não Recomenda</th>
            <th class="text-end">Total</th>
          </tr>
        </thead>
        <tbody id="pessoas-tbody">
          {% for p in data.pessoas_ranking %}
          <tr class="pessoa-row" data-nome="{{ p.nome|lower }}">
            <td class="text-muted">{{ loop.index }}</td>
            <td>{{ p.nome }}</td>
            <td class="text-end">
              <span class="badge {% if p.pct >= 80 %}bg-success{% elif p.pct >= 50 %}bg-warning{% else %}bg-danger{% endif %}">
                {{ "%.0f"|format(p.pct) }}%
              </span>
            </td>
            <td class="text-end">{{ p.rec }}</td>
            <td class="text-end">{{ p.nao_rec }}</td>
            <td class="text-end">{{ p.total }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

  <!-- AÇÕES -->
  <div class="tab-content hidden" id="tab-acoes">
    {% set areas_criticas = data.areas | selectattr('zona', 'equalto', 'critica') | list %}
    {% set areas_atencao = data.areas | selectattr('zona', 'equalto', 'atencao') | list %}
    {% set areas_excelencia = data.areas | selectattr('zona', 'equalto', 'excelencia') | list %}

    {% if areas_criticas %}
    <div class="action-card action-critica">
      <h4>🔴 Zona Crítica — Ação imediata</h4>
      <p class="text-muted">Áreas com NPS negativo precisam de intervenção estruturada:</p>
      <ul>
        {% for a in areas_criticas | sort(attribute='nps') %}
        <li><strong>{{ a.nome }}</strong> (NPS {{ "%.1f"|format(a.nps) }}) — conversa 1:1 com liderança + 1-2 ações concretas em 30 dias + re-medir em 90 dias</li>
        {% endfor %}
      </ul>
    </div>
    {% endif %}

    {% if areas_atencao %}
    <div class="action-card action-atencao">
      <h4>🟡 Zona de Atenção — Priorizar melhorias</h4>
      <p class="text-muted">Começar pelas áreas com mais detratores e comentários concretos:</p>
      <ul>
        {% for a in areas_atencao | sort(attribute='nps') %}
        <li><strong>{{ a.nome }}</strong> (NPS {{ "%.1f"|format(a.nps) }}) — {{ a.detratores }} detratores em {{ a.total }} avaliações</li>
        {% endfor %}
      </ul>
    </div>
    {% endif %}

    {% if areas_excelencia %}
    <div class="action-card action-excelencia">
      <h4>🟢 Zona de Excelência — Proteger e escalar</h4>
      <p class="text-muted">Mapear o que essas áreas fazem diferente e usar como modelo:</p>
      <ul>
        {% for a in areas_excelencia | sort(attribute='nps', reverse=true) %}
        <li><strong>{{ a.nome }}</strong> (NPS {{ "%.1f"|format(a.nps) }})</li>
        {% endfor %}
      </ul>
    </div>
    {% endif %}

    {% if data.alertas %}
    <div class="action-card" style="border-left-color: #6f42c1;">
      <h4>👤 Pessoas em Alerta</h4>
      <p class="text-muted">Conversas 1:1 com cada pessoa abaixo dentro de 30 dias. Cruzar com feedback 360 existente.</p>
      <ul>
        {% for p in data.alertas %}
        <li><strong>{{ p.nome }}</strong> — {{ "%.0f"|format(p.pct) }}% de recomendação ({{ p.nao_rec }}/{{ p.total }} não recomendam)</li>
        {% endfor %}
      </ul>
    </div>
    {% endif %}

    {% set big_gaps = data.divergencias | selectattr('gap', 'gt', 15) | list + data.divergencias | selectattr('gap', 'lt', -15) | list %}
    {% if big_gaps %}
    <div class="action-card" style="border-left-color: #0dcaf0;">
      <h4>🔀 Calibração Liderança x IC</h4>
      <p class="text-muted">Áreas com gap &gt; 15 pontos — conversa direta pra alinhar percepção:</p>
      <ul>
        {% for d in big_gaps %}
        <li><strong>{{ d.area }}</strong> — gap {{ "%+.1f"|format(d.gap) }} ({% if d.gap > 0 %}Liderança mais otimista{% else %}IC mais otimista{% endif %})</li>
        {% endfor %}
      </ul>
    </div>
    {% endif %}
  </div>

  <div class="text-center footer-info mt-4">
    Dashboard Pesquisa G4 · Dados em tempo real do SQLite
  </div>
</div>

<script>
  document.querySelectorAll('#tabs .nav-link').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      document.querySelectorAll('#tabs .nav-link').forEach(l => l.classList.remove('active'));
      link.classList.add('active');
      document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
      document.getElementById('tab-' + link.dataset.tab).classList.remove('hidden');
    });
  });

  function toggleBreakdown(idx) {
    const row = document.querySelector(`.ranking-row[data-idx="${idx}"]`);
    const breakdown = document.getElementById('breakdown-' + idx);
    if (!row || !breakdown) return;
    row.classList.toggle('active');
    breakdown.classList.toggle('hidden');
  }

  function toggleArea(card) {
    card.classList.toggle('expanded');
    card.querySelector('.area-details').classList.toggle('hidden');
    const toggle = card.querySelector('.detail-toggle');
    if (toggle.textContent.startsWith('▼')) toggle.textContent = toggle.textContent.replace('▼', '▲');
    else toggle.textContent = toggle.textContent.replace('▲', '▼');
  }

  document.getElementById('search-ranking').addEventListener('input', (e) => {
    const q = e.target.value.toLowerCase();
    document.querySelectorAll('#ranking-tbody tr.ranking-row').forEach(tr => {
      const visible = tr.dataset.nome.includes(q);
      tr.style.display = visible ? '' : 'none';
      const idx = tr.dataset.idx;
      const bd = document.getElementById('breakdown-' + idx);
      if (bd) {
        bd.classList.add('hidden');
        tr.classList.remove('active');
      }
    });
  });

  document.querySelectorAll('#tab-ranking .filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const filter = btn.dataset.filter;
      document.querySelectorAll('#tab-ranking .filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      document.querySelectorAll('#ranking-tbody tr.ranking-row').forEach(tr => {
        const visible = (filter === 'todos' || tr.dataset.zona === filter);
        tr.style.display = visible ? '' : 'none';
        const idx = tr.dataset.idx;
        const bd = document.getElementById('breakdown-' + idx);
        if (bd) {
          bd.classList.add('hidden');
          tr.classList.remove('active');
        }
      });
    });
  });

  document.getElementById('search-areas').addEventListener('input', (e) => {
    const q = e.target.value.toLowerCase();
    document.querySelectorAll('#areas-container .area-card').forEach(card => {
      card.style.display = card.dataset.nome.includes(q) ? '' : 'none';
    });
  });

  document.getElementById('search-pessoas').addEventListener('input', (e) => {
    const q = e.target.value.toLowerCase();
    document.querySelectorAll('#pessoas-tbody tr').forEach(tr => {
      tr.style.display = tr.dataset.nome.includes(q) ? '' : 'none';
    });
  });
</script>
</body>
</html>
"""
