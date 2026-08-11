<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<title>GRAUS Fleet Dashboard</title>

<!-- Leaflet for the map -->
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>

<link rel="stylesheet" href="style.css?v=2" />
</head>
<body>

  <div id="graus-fleet-addin">

    <header class="gf-header">
      <h1>Cruscotto Flotta GRAUS</h1>
      <div class="gf-header-actions">
        <select id="gf-vehicle-select" class="gf-select">
          <option value="">Tutta la flotta</option>
        </select>
        <button id="gf-refresh" class="gf-btn">Aggiorna</button>
      </div>
    </header>

    <section class="gf-kpis" id="gf-kpis">
      <div class="gf-kpi">
        <span class="gf-kpi-value" id="kpi-total">–</span>
        <span class="gf-kpi-label">Veicoli totali</span>
      </div>
      <div class="gf-kpi">
        <span class="gf-kpi-value" id="kpi-moving">–</span>
        <span class="gf-kpi-label">In movimento</span>
      </div>
      <div class="gf-kpi">
        <span class="gf-kpi-value" id="kpi-stopped">–</span>
        <span class="gf-kpi-label">Fermi / in pausa</span>
      </div>
      <div class="gf-kpi">
        <span class="gf-kpi-value" id="kpi-offline">–</span>
        <span class="gf-kpi-label">Offline (no segnale)</span>
      </div>
    </section>

    <main class="gf-main">
      <div id="gf-map"></div>

      <aside class="gf-sidebar">
        <div id="gf-fleet-panel">
          <h2>Fermate / Pause in corso</h2>
          <div id="gf-stops-list" class="gf-stops-list">
            <p class="gf-empty">Caricamento…</p>
          </div>
        </div>

        <div id="gf-vehicle-panel" class="gf-hidden">
          <h2 id="gf-vehicle-name">–</h2>
          <div class="gf-vehicle-meta" id="gf-vehicle-meta"></div>
          <h3>Soste di oggi</h3>
          <div id="gf-vehicle-stops" class="gf-stops-list"></div>
        </div>
      </aside>
    </main>

  </div>

  <script src="main.js?v=3"></script>
</body>
</html>
