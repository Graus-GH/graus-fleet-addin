/*
 * GRAUS Fleet Dashboard — Geotab Map Add-In
 *
 * Map Add-Ins (config with "page": "map") use a DIFFERENT contract than
 * standard navigation Add-Ins:
 *
 *   geotab.addin.<uniqueName> = (elt, service) => {
 *     // elt     = the HTMLElement container for this Add-In's UI
 *     // service = { page, api, localStorage, events, map, tooltip, actionList, canvas }
 *   };
 *
 * This function is called once when the user visits the Map/Trips History
 * page. Subsequent focus/blur happen via service.page.attach('focus', ...)
 * and service.page.attach('blur', ...) — NOT via a returned {focus, blur}
 * object like standard nav Add-Ins.
 *
 * service.api.call(method, params) returns a Promise directly (no callback).
 */
geotab.addin.grausFleetDashboard = (elt, service) => {

  let map;
  let markersLayer;
  let routeLayer;
  let refreshTimer;

  let selectedDeviceId = null;   // null = fleet view, otherwise a single vehicle
  let allTripsToday = [];        // cached raw trips, reused for the vehicle detail panel
  let vehicleListPopulated = false;

  const REFRESH_INTERVAL_MS = 60 * 1000; // 1 min live refresh
  const OFFLINE_THRESHOLD_MS = 15 * 60 * 1000; // no update in 15 min = offline

  function initMap() {
    map = L.map("gf-map").setView([46.55, 11.9], 11); // default: Alta Badia area
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "© OpenStreetMap contributors",
      maxZoom: 18
    }).addTo(map);
    markersLayer = L.layerGroup().addTo(map);
    routeLayer = L.layerGroup().addTo(map);
  }

  function fmtDuration(ms) {
    const totalMin = Math.max(0, Math.round(ms / 60000));
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    return h > 0 ? `${h}h ${m}m` : `${m}m`;
  }

  // Map Add-In api service: call(method, params) returns a Promise directly
  function apiCall(method, params) {
    return service.api.call(method, params);
  }

  async function loadFleetData() {
    const now = new Date();
    const startOfToday = new Date(now);
    startOfToday.setHours(0, 0, 0, 0);

    // 1. All active devices (vehicles)
    const devices = await apiCall("Get", {
      typeName: "Device",
      search: { fromDate: now.toISOString() }
    });

    // 2. Live status snapshot for every device (position, speed, isDriving)
    const statuses = await apiCall("Get", { typeName: "DeviceStatusInfo" });
    const statusByDevice = {};
    statuses.forEach(s => { statusByDevice[s.device.id] = s; });

    // 3. Today's trips per device, used to estimate current stop/pause duration.
    //    NOTE: this is an approximation (time since the last trip ended).
    //    For production-grade accuracy, replace with an ExceptionEvent-based
    //    "Stop" Rule in MyGeotab and query ExceptionEvent instead.
    const trips = await apiCall("Get", {
      typeName: "Trip",
      search: { fromDate: startOfToday.toISOString() }
    });
    allTripsToday = trips; // cache for the vehicle detail panel

    const lastTripByDevice = {};
    trips.forEach(t => {
      const id = t.device.id;
      if (!lastTripByDevice[id] || new Date(t.stop) > new Date(lastTripByDevice[id].stop)) {
        lastTripByDevice[id] = t;
      }
    });

    return devices.map(device => {
      const status = statusByDevice[device.id];
      const lastTrip = lastTripByDevice[device.id];

      let state = "offline";
      let stopSince = null;

      if (status) {
        const lastUpdateAge = now - new Date(status.dateTime);
        if (lastUpdateAge > OFFLINE_THRESHOLD_MS) {
          state = "offline";
        } else if (status.isDriving) {
          state = "moving";
        } else {
          state = "stopped";
          stopSince = lastTrip ? new Date(lastTrip.stop) : new Date(status.dateTime);
        }
      }

      return {
        id: device.id,
        name: device.name,
        latitude: status ? status.latitude : null,
        longitude: status ? status.longitude : null,
        speed: status ? status.speed : 0,
        state,
        stopDurationMs: stopSince ? now - stopSince : null
      };
    });
  }

  function renderKpis(vehicles) {
    document.getElementById("kpi-total").textContent = vehicles.length;
    document.getElementById("kpi-moving").textContent =
      vehicles.filter(v => v.state === "moving").length;
    document.getElementById("kpi-stopped").textContent =
      vehicles.filter(v => v.state === "stopped").length;
    document.getElementById("kpi-offline").textContent =
      vehicles.filter(v => v.state === "offline").length;
  }

  function renderMap(vehicles) {
    markersLayer.clearLayers();
    const shown = selectedDeviceId
      ? vehicles.filter(v => v.id === selectedDeviceId)
      : vehicles;
    const withPosition = shown.filter(v => v.latitude && v.longitude);

    withPosition.forEach(v => {
      const color = v.state === "moving" ? "#2e7d32"
                  : v.state === "stopped" ? "#c62828"
                  : "#9e9e9e";
      const marker = L.circleMarker([v.latitude, v.longitude], {
        radius: selectedDeviceId ? 9 : 7,
        color,
        fillColor: color,
        fillOpacity: 0.9
      }).bindTooltip(v.name, { permanent: false, className: "gf-marker-label" });
      // Clicking a marker in fleet view selects that vehicle
      marker.on("click", () => {
        if (!selectedDeviceId) {
          document.getElementById("gf-vehicle-select").value = v.id;
          document.getElementById("gf-vehicle-select").dispatchEvent(new Event("change"));
        }
      });
      marker.addTo(markersLayer);
    });

    if (!selectedDeviceId && withPosition.length) {
      const bounds = L.latLngBounds(withPosition.map(v => [v.latitude, v.longitude]));
      map.fitBounds(bounds.pad(0.2));
    } else if (selectedDeviceId && withPosition.length === 1) {
      map.setView([withPosition[0].latitude, withPosition[0].longitude], 14);
    }
  }

  // Draws the day's route (breadcrumb trail) for a single selected vehicle
  async function renderVehicleRoute(deviceId) {
    routeLayer.clearLayers();
    if (!deviceId) return;

    const now = new Date();
    const startOfToday = new Date(now);
    startOfToday.setHours(0, 0, 0, 0);

    const points = await apiCall("Get", {
      typeName: "LogRecord",
      search: {
        deviceSearch: { id: deviceId },
        fromDate: startOfToday.toISOString(),
        toDate: now.toISOString()
      },
      resultsLimit: 5000
    });

    if (!points.length) return;

    points.sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));
    const latlngs = points
      .filter(p => p.latitude && p.longitude)
      .map(p => [p.latitude, p.longitude]);

    if (latlngs.length > 1) {
      L.polyline(latlngs, { color: "#005caa", weight: 3, opacity: 0.75 }).addTo(routeLayer);
      map.fitBounds(L.latLngBounds(latlngs).pad(0.15));
    }
  }

  function populateVehicleSelect(vehicles) {
    if (vehicleListPopulated) return; // build the option list only once
    const select = document.getElementById("gf-vehicle-select");
    vehicles
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach(v => {
        const opt = document.createElement("option");
        opt.value = v.id;
        opt.textContent = v.name;
        select.appendChild(opt);
      });
    vehicleListPopulated = true;
  }

  function renderVehiclePanel(vehicle) {
    document.getElementById("gf-fleet-panel").classList.add("gf-hidden");
    document.getElementById("gf-vehicle-panel").classList.remove("gf-hidden");

    document.getElementById("gf-vehicle-name").textContent = vehicle ? vehicle.name : "–";

    const stateLabel = { moving: "In movimento", stopped: "Fermo", offline: "Offline" };
    const meta = document.getElementById("gf-vehicle-meta");
    if (vehicle) {
      meta.innerHTML = `
        <div>Stato: <strong>${stateLabel[vehicle.state] || "–"}</strong></div>
        <div>Velocità: <strong>${Math.round(vehicle.speed || 0)} km/h</strong></div>
        ${vehicle.state === "stopped" && vehicle.stopDurationMs
          ? `<div>Fermo da: <strong>${fmtDuration(vehicle.stopDurationMs)}</strong></div>`
          : ""}
      `;
    } else {
      meta.innerHTML = "";
    }

    const deviceTrips = allTripsToday
      .filter(t => t.device.id === (vehicle ? vehicle.id : null))
      .sort((a, b) => new Date(b.stop) - new Date(a.stop));

    const stopsContainer = document.getElementById("gf-vehicle-stops");
    if (!deviceTrips.length) {
      stopsContainer.innerHTML = '<p class="gf-empty">Nessun viaggio registrato oggi.</p>';
      return;
    }

    stopsContainer.innerHTML = deviceTrips.map(t => {
      const stopTime = new Date(t.stop).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
      const startTime = new Date(t.start).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
      const distanceKm = t.distance ? (t.distance).toFixed(1) : "0.0";
      return `
        <div class="gf-stop-item">
          <div class="gf-device-name">${startTime} → ${stopTime}</div>
          <div>${distanceKm} km percorsi</div>
          ${t.stopDuration ? `<div class="gf-duration">Sosta successiva: ${fmtDuration(t.stopDuration * 1000)}</div>` : ""}
        </div>
      `;
    }).join("");
  }

  function showFleetPanel() {
    document.getElementById("gf-vehicle-panel").classList.add("gf-hidden");
    document.getElementById("gf-fleet-panel").classList.remove("gf-hidden");
  }

  function renderStopsList(vehicles) {
    const container = document.getElementById("gf-stops-list");
    const stopped = vehicles
      .filter(v => v.state === "stopped")
      .sort((a, b) => (b.stopDurationMs || 0) - (a.stopDurationMs || 0));

    if (!stopped.length) {
      container.innerHTML = '<p class="gf-empty">Nessun veicolo fermo al momento.</p>';
      return;
    }

    container.innerHTML = stopped.map(v => `
      <div class="gf-stop-item">
        <div class="gf-device-name">${v.name}</div>
        <div class="gf-duration">Fermo da ${fmtDuration(v.stopDurationMs || 0)}</div>
      </div>
    `).join("");
  }

  async function refresh() {
    try {
      const vehicles = await loadFleetData();
      populateVehicleSelect(vehicles);
      renderKpis(vehicles);
      renderMap(vehicles);

      if (selectedDeviceId) {
        const vehicle = vehicles.find(v => v.id === selectedDeviceId);
        renderVehiclePanel(vehicle);
        await renderVehicleRoute(selectedDeviceId);
      } else {
        showFleetPanel();
        renderStopsList(vehicles);
      }
    } catch (err) {
      console.error("GRAUS Fleet Dashboard — errore caricamento dati:", err);
      document.getElementById("gf-stops-list").innerHTML =
        '<p class="gf-empty">Errore nel caricamento dei dati. Riprova.</p>';
    }
  }

  // --- Entry point: this code runs once, when the user visits the page ---

  initMap();

  document.getElementById("gf-refresh").addEventListener("click", () => refresh());

  document.getElementById("gf-vehicle-select").addEventListener("change", (e) => {
    selectedDeviceId = e.target.value || null;
    routeLayer.clearLayers();
    refresh();
  });

  refresh();
  refreshTimer = setInterval(() => refresh(), REFRESH_INTERVAL_MS);

  // Map Add-In focus/blur come through the page service, not a returned object
  service.page.attach("focus", () => {
    refresh();
    if (!refreshTimer) {
      refreshTimer = setInterval(() => refresh(), REFRESH_INTERVAL_MS);
    }
  });

  service.page.attach("blur", () => {
    if (refreshTimer) {
      clearInterval(refreshTimer);
      refreshTimer = null;
    }
  });

};
