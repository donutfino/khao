// แอพพลิเคชันควบคุมการทำงานหลัก (Main Application Controller - ฉบับแก้ไขใหม่)
document.addEventListener('DOMContentLoaded', function() {
  
  // ==========================================
  // 1. สถานะแอปพลิเคชัน (Application State)
  // ==========================================
  let appState = {
    importedData: [],       // รายการข้อมูลดิบที่อิมพอร์ตเข้ามา
    customers: [],          // รายการลูกค้า (คัดแยกคลังสินค้าออกแล้ว)
    depot: null,            // วัตถุจุดคลังสินค้าเริ่มต้น
    solvedResult: null,     // ผลการจัดเส้นทางจาก solver
    map: null,              // ตัวแปรแผนที่ Leaflet
    markersLayer: null,     // Layer Group สำหรับหมุดบนแผนที่
    routesLayer: null,      // Layer Group สำหรับเส้นทางบนแผนที่
    mapMarkers: {},         // เก็บ Object หมุดแผนที่แยกตาม Customer ID เพื่อเรียกไฮไลท์
    fileName: ''            // ชื่อไฟล์ที่โหลดเข้ามา
  };

  // ==========================================
  // 2. การเตรียมความพร้อมเบื้องต้น (Init Map & Elements)
  // ==========================================
  function initApp() {
    initMap();
    initTabEvents();
    initUploadEvents();
    initSliderEvents();
    initActionEvents();
    initDatabase(); // เริ่มต้นระบบฐานข้อมูลในเบราว์เซอร์
    initTemplateDownload(); // สร้างเทมเพลต Excel สำหรับปุ่มดาวน์โหลด
  }

  // เริ่มต้นสร้างแผนที่ Leaflet
  function initMap() {
    const bangkokCoords = [13.7563, 100.5018];
    
    appState.map = L.map('map', {
      zoomControl: true,
      scrollWheelZoom: true
    }).setView(bangkokCoords, 11);

    // ใช้แผนที่สีเข้มพรีเมียม (CartoDB Dark Matter)
    L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
      subdomains: 'abcd',
      maxZoom: 20
    }).addTo(appState.map);

    appState.markersLayer = L.featureGroup().addTo(appState.map);
    appState.routesLayer = L.featureGroup().addTo(appState.map);

    // ดักจับการคลิกแผนที่เพื่อป้อนพิกัดคลังแบบแมนนวล
    appState.map.on('click', function(e) {
      const customBtn = document.getElementById('depot-source-custom');
      const isCustomActive = customBtn && customBtn.classList.contains('active');
      if (isCustomActive) {
        const { lat, lng } = e.latlng;
        document.getElementById('custom-depot-lat').value = lat.toFixed(6);
        document.getElementById('custom-depot-lng').value = lng.toFixed(6);
        
        updateDepotAndCustomers();
        plotPointsBeforeSolve();
        showToast(`เลือกพิกัดคลังสำเร็จ: ${lat.toFixed(4)}, ${lng.toFixed(4)}`, 'success');
      }
    });
  }

  // ==========================================
  // 3. แท็บเมนู (Tabs Handling)
  // ==========================================
  function initTabEvents() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
      btn.addEventListener('click', function() {
        const targetTabId = this.getAttribute('data-tab');
        
        tabBtns.forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        
        const panes = document.querySelectorAll('.tab-pane');
        panes.forEach(pane => {
          pane.classList.remove('active');
          if (pane.id === targetTabId) {
            pane.classList.add('active');
          }
        });

        setTimeout(() => {
          appState.map.invalidateSize();
        }, 100);
      });
    });
  }

  // ==========================================
  // 4. นำเข้า Excel/CSV และแปลงข้อมูล
  // ==========================================
  function initUploadEvents() {
    const dropzone = document.getElementById('excel-dropzone');
    const fileInput = document.getElementById('file-input');
    const btnRemove = document.getElementById('btn-remove-file');

    dropzone.addEventListener('click', () => fileInput.click());

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        processUploadedFile(files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      const files = e.target.files;
      if (files.length > 0) {
        processUploadedFile(files[0]);
      }
    });

    btnRemove.addEventListener('click', (e) => {
      e.stopPropagation();
      resetImportState();
      showToast('นำข้อมูลออกเรียบร้อยแล้ว', 'info');
    });
  }

  function processUploadedFile(file) {
    appState.fileName = file.name;
    const reader = new FileReader();

    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rawJson = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: "" });
        
        if (rawJson.length === 0) {
          throw new Error('ไม่พบข้อมูลจุดจอดในไฟล์นี้');
        }

        parseExcelData(rawJson);
      } catch (err) {
        showToast('ผิดพลาดในการอ่านไฟล์: ' + err.message, 'error');
        console.error(err);
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function parseExcelData(rawData) {
    const parsed = [];
    
    rawData.forEach((row, idx) => {
      const name = findValueByKeys(row, ['customer_name', 'name', 'customer', 'ชื่อ', 'ชื่อลูกค้า', 'จุดส่ง']);
      const address = findValueByKeys(row, ['address', 'ที่อยู่', 'รายละเอียด']);
      const rawLat = findValueByKeys(row, ['latitude', 'lat', 'y', 'ละติจูด']);
      const rawLng = findValueByKeys(row, ['longitude', 'lng', 'lon', 'x', 'ลองจิจูด']);
      const route = findValueByKeys(row, ['route', 'vehicle', 'สาย', 'สายส่ง', 'สายส่งสินค้า']);

      if (rawLat !== null && rawLng !== null) {
        // ทำความสะอาดค่าพิกัด (เช่น แปลงเครื่องหมายจุลภาคเป็นจุด)
        const latStr = String(rawLat).replace(/,/g, '.').trim();
        const lngStr = String(rawLng).replace(/,/g, '.').trim();

        let lat = parseFloat(latStr);
        let lng = parseFloat(lngStr);

        if (!isNaN(lat) && !isNaN(lng)) {
          // ตรวจสอบความถูกต้องและสลับค่าพิกัดอัตโนมัติหากใส่ละติจูด/ลองจิจูดสลับช่องกัน
          // โดยปกติละติจูดประเทศไทยจะอยู่ระหว่าง 5.5 ถึง 20.5 และลองจิจูดอยู่ระหว่าง 97.2 ถึง 105.7
          if ((lat >= 97.2 && lat <= 105.7) && (lng >= 5.5 && lng <= 20.5)) {
            const temp = lat;
            lat = lng;
            lng = temp;
          }

          parsed.push({
            id: idx + 1,
            name: name ? String(name).trim() : `จุดส่งสินค้าที่ ${idx + 1}`,
            address: address ? String(address).trim() : `พิกัด: ${lat.toFixed(4)}, ${lng.toFixed(4)}`,
            lat: lat,
            lng: lng,
            route: route ? String(route).trim() : undefined,
            isDepot: false
          });
        }
      }
    });

    if (parsed.length === 0) {
      showToast('ไม่พบแถวข้อมูลที่มีละติจูด/ลองจิจูดครบถ้วน', 'error');
      return;
    }

    appState.importedData = parsed;
    appState.importedData[0].isDepot = true;
    
    updateDataUIs();
    showToast(`อิมพอร์ตข้อมูลสำเร็จ ${parsed.length} จุดส่ง`, 'success');
  }

  function findValueByKeys(row, keys) {
    const normalize = str => {
      if (typeof str !== 'string') return '';
      return str.toLowerCase().replace(/[^a-z0-9\u0e00-\u0e7f]/ig, '');
    };

    // ขั้นตอนที่ 1: ตรวจสอบแบบจับคู่ตรงกันหลังจาก normalize (Exact normalized match)
    for (let key of keys) {
      const normalizedKey = normalize(key);
      if (!normalizedKey) continue;
      
      const matchedKey = Object.keys(row).find(k => normalize(k) === normalizedKey);
      if (matchedKey !== undefined && row[matchedKey] !== null && row[matchedKey] !== '') {
        return row[matchedKey];
      }
    }

    // ขั้นตอนที่ 2: ตรวจสอบแบบข้อความย่อย (Substring match) เผื่อหัวตารางมีวงเล็บ เช่น "ชื่อลูกค้า (Customer Name)"
    for (let key of keys) {
      const normalizedKey = normalize(key);
      if (!normalizedKey) continue;
      
      const matchedKey = Object.keys(row).find(k => {
        const normK = normalize(k);
        return normK && (normK.includes(normalizedKey) || normalizedKey.includes(normK));
      });
      
      if (matchedKey !== undefined && row[matchedKey] !== null && row[matchedKey] !== '') {
        return row[matchedKey];
      }
    }
    return null;
  }

  function resetImportState() {
    appState.importedData = [];
    appState.customers = [];
    appState.depot = null;
    appState.solvedResult = null;
    appState.fileName = '';
    
    document.getElementById('file-input').value = '';
    document.getElementById('file-info-box').style.display = 'none';
    document.getElementById('preview-card').style.display = 'none';
    if (document.getElementById('save-location-card')) {
      document.getElementById('save-location-card').style.display = 'none';
    }
    
    document.getElementById('btn-solve-routing').setAttribute('disabled', 'true');
    
    document.getElementById('results-active-view').style.display = 'none';
    document.getElementById('results-empty-state').style.display = 'flex';
    
    appState.markersLayer.clearLayers();
    appState.routesLayer.clearLayers();
    appState.map.setView([13.7563, 100.5018], 11);
  }

  // ==========================================
  // 5. โหลด Demo Data และดาวน์โหลด Template
  // ==========================================
  function initActionEvents() {
    document.getElementById('btn-load-sample').addEventListener('click', () => {
      appState.importedData = JSON.parse(JSON.stringify(window.SAMPLE_DATA));
      appState.fileName = 'ข้อมูลทดสอบการส่งสินค้ากทม.xlsx';
      
      // ล้างคอลัมน์ปริมาณสินค้าจากข้อมูลตัวอย่าง
      appState.importedData.forEach(item => {
        delete item.demand;
      });

      updateDataUIs();
      showToast('โหลดข้อมูลตัวอย่างกรุงเทพฯ 16 จุด เรียบร้อย', 'success');
    });

    document.getElementById('btn-download-template').addEventListener('click', () => {
      showToast('ดาวน์โหลดเทมเพลตเรียบร้อยแล้ว', 'success');
    });

    document.getElementById('btn-solve-routing').addEventListener('click', () => {
      solveRoutingAndDisplay();
    });

    document.getElementById('btn-export-excel').addEventListener('click', () => {
      exportResultsToExcel();
    });

    // ปุ่มสลับแหล่งที่มาคลังสินค้า (ตารางจุดส่ง vs กำหนดเอง)
    document.getElementById('depot-source-import').addEventListener('click', function() {
      this.classList.add('active');
      document.getElementById('depot-source-custom').classList.remove('active');
      document.getElementById('depot-import-wrapper').style.display = 'block';
      document.getElementById('depot-custom-wrapper').style.display = 'none';
      
      updateDepotAndCustomers();
      plotPointsBeforeSolve();
      showToast('สลับเป็นแบบเลือกจากตารางข้อมูลจุดจอด', 'info');
    });

    document.getElementById('depot-source-custom').addEventListener('click', function() {
      this.classList.add('active');
      document.getElementById('depot-source-import').classList.remove('active');
      document.getElementById('depot-custom-wrapper').style.display = 'flex';
      document.getElementById('depot-import-wrapper').style.display = 'none';

      // เติมค่าพิกัดเริ่มต้นถ้ายังว่างอยู่
      const latInput = document.getElementById('custom-depot-lat');
      const lngInput = document.getElementById('custom-depot-lng');
      if (!latInput.value && appState.depot) {
        latInput.value = appState.depot.lat.toFixed(6);
        lngInput.value = appState.depot.lng.toFixed(6);
      } else if (!latInput.value) {
        latInput.value = "13.7762";
        lngInput.value = "100.5748";
      }

      updateDepotAndCustomers();
      plotPointsBeforeSolve();
      showToast('ระบุพิกัดคลังเอง: สามารถคลิกเลือกจุดบนแผนที่ได้โดยตรง', 'info');
    });

    // ดักจับเหตุการณ์แก้ไขพิกัดหรือชื่อคลังแมนนวล
    ['custom-depot-name', 'custom-depot-lat', 'custom-depot-lng'].forEach(id => {
      document.getElementById(id).addEventListener('input', () => {
        const isCustomActive = document.getElementById('depot-source-custom').classList.contains('active');
        if (isCustomActive) {
          updateDepotAndCustomers();
          plotPointsBeforeSolve();
        }
      });
    });
  }

  function updateDataUIs() {
    const infoBox = document.getElementById('file-info-box');
    const nameEl = document.getElementById('file-name');
    const statsEl = document.getElementById('file-stats');
    
    infoBox.style.display = 'flex';
    nameEl.textContent = appState.fileName;
    statsEl.textContent = `พบข้อมูลทั้งหมด ${appState.importedData.length} จุด (รวม Depot)`;

    const selectDepot = document.getElementById('select-depot');
    selectDepot.innerHTML = '';
    
    appState.importedData.forEach((item, idx) => {
      const option = document.createElement('option');
      option.value = item.id;
      option.textContent = `${idx + 1}. ${item.name} (${item.lat.toFixed(4)}, ${item.lng.toFixed(4)})`;
      if (item.isDepot) {
        option.selected = true;
      }
      selectDepot.appendChild(option);
    });

    updateDepotAndCustomers();

    const previewCard = document.getElementById('preview-card');
    const tbody = document.querySelector('#data-preview-table tbody');
    tbody.innerHTML = '';

    appState.importedData.forEach(item => {
      const tr = document.createElement('tr');
      if (item.isDepot) {
        tr.style.background = 'rgba(239, 68, 68, 0.08)';
      }
      tr.innerHTML = `
        <td title="${item.name}">${item.isDepot ? '📦 [คลัง] ' : ''}${item.name}</td>
        <td>${item.isDepot ? '-' : (item.route || '<span style="color:var(--text-muted)">จัดสรรอัตโนมัติ</span>')}</td>
        <td title="${item.address}">${item.isDepot ? '-' : item.address}</td>
        <td>${item.lat.toFixed(4)}</td>
        <td>${item.lng.toFixed(4)}</td>
      `;
      tbody.appendChild(tr);
    });
    previewCard.style.display = 'block';
    if (document.getElementById('save-location-card')) {
      document.getElementById('save-location-card').style.display = 'block';
    }

    document.getElementById('btn-solve-routing').removeAttribute('disabled');

    plotPointsBeforeSolve();
  }

  function updateDepotAndCustomers() {
    const customBtn = document.getElementById('depot-source-custom');
    const isCustomActive = customBtn && customBtn.classList.contains('active');
    
    if (isCustomActive) {
      const name = document.getElementById('custom-depot-name').value || 'คลังสินค้ากำหนดเอง';
      const latVal = document.getElementById('custom-depot-lat').value;
      const lngVal = document.getElementById('custom-depot-lng').value;
      
      const latStr = String(latVal || '').replace(/,/g, '.').trim();
      const lngStr = String(lngVal || '').replace(/,/g, '.').trim();

      let lat = parseFloat(latStr);
      let lng = parseFloat(lngStr);

      if (isNaN(lat) || isNaN(lng)) {
        lat = 13.7762;
        lng = 100.5748;
      } else {
        // ตรวจสอบและสลับค่าพิกัดอัตโนมัติหากใส่ละติจูด/ลองจิจูดสลับช่องกัน
        if ((lat >= 97.2 && lat <= 105.7) && (lng >= 5.5 && lng <= 20.5)) {
          const temp = lat;
          lat = lng;
          lng = temp;
        }
      }

      appState.depot = {
        id: 'custom-depot',
        name: name,
        address: `คลังกำหนดเอง (${lat.toFixed(6)}, ${lng.toFixed(6)})`,
        lat: lat,
        lng: lng,
        isDepot: true
      };

      // ทุกจุดส่งที่โหลดเข้ามาจะเป็นลูกค้าปลายทาง (Customers) ทั้งหมด เพราะคลังตั้งอยู่ภายนอกแบบระบุพิกัดแยกต่างหาก
      appState.customers = appState.importedData.map(item => {
        const copyItem = {...item};
        copyItem.isDepot = false;
        return copyItem;
      });
    } else {
      const selectDepot = document.getElementById('select-depot');
      if (selectDepot && selectDepot.value) {
        const selectedId = parseInt(selectDepot.value);
        
        appState.importedData.forEach(item => {
          if (item.id === selectedId) {
            item.isDepot = true;
            appState.depot = item;
          } else {
            item.isDepot = false;
          }
        });
      }
      appState.customers = appState.importedData.filter(item => !item.isDepot);
    }
  }

  // ดักการเปลี่ยน Depot จาก dropdown
  const selectDepotEl = document.getElementById('select-depot');
  if (selectDepotEl) {
    selectDepotEl.addEventListener('change', () => {
      updateDepotAndCustomers();
      plotPointsBeforeSolve(); // อัปเดตพอยต์คลังบนแผนที่
      showToast('เปลี่ยนจุดคลังสินค้าเริ่มต้นเรียบร้อยแล้ว', 'info');
    });
  }

  function plotPointsBeforeSolve() {
    appState.markersLayer.clearLayers();
    appState.routesLayer.clearLayers();
    appState.mapMarkers = {};

    if (appState.depot) {
      const depotIcon = L.divIcon({
        className: 'custom-div-icon',
        html: `<div class="depot-marker-pin">
                <svg viewBox="0 0 24 24"><path d="M19 4H5c-1.11 0-2 .9-2 2v12c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm-9 14H5v-5h5v5zm0-6H5V7h5v5zm9 6h-5v-5h5v5zm0-6h-5V7h5v5z"/></svg>
              </div>`,
        iconSize: [28, 28],
        iconAnchor: [14, 14]
      });

      const depotMarker = L.marker([appState.depot.lat, appState.depot.lng], { icon: depotIcon })
        .bindPopup(`<div class="popup-title">${appState.depot.name}</div><div class="popup-detail">จุดเริ่มต้นและสิ้นสุดเดินรถ (Depot)</div><div class="popup-detail">${appState.depot.address}</div>`)
        .addTo(appState.markersLayer);
      
      appState.mapMarkers[appState.depot.id] = depotMarker;
    }

    appState.customers.forEach(cust => {
      const custIcon = L.divIcon({
        className: 'custom-div-icon',
        html: `<div class="map-marker-pin" style="background-color: #6b7280;"></div>`,
        iconSize: [24, 24],
        iconAnchor: [12, 12]
      });

      const marker = L.marker([cust.lat, cust.lng], { icon: custIcon })
        .bindPopup(`
          <div class="popup-title">${cust.name}</div>
          <div class="popup-detail">${cust.address}</div>
        `)
        .addTo(appState.markersLayer);
      
      appState.mapMarkers[cust.id] = marker;
    });

    if (appState.importedData.length > 0) {
      appState.map.fitBounds(appState.markersLayer.getBounds(), { padding: [40, 40] });
    }
  }

  function initTemplateDownload() {
    const wsData = [
      ["Customer_Name", "Address", "Latitude", "Longitude", "Route"],
      ["คลังสินค้าหลัก (Depot)", "ศูนย์กระจายสินค้าดินแดง กรุงเทพฯ", 13.7700, 100.5500, ""],
      ["ร้านค้าลาดพร้าว 101", "ซอยลาดพร้าว 101 วังทองหลาง กรุงเทพฯ", 13.7954, 100.6288, "สายที่ 1"],
      ["ห้างโลตัส บางกะปิ", "ถนนลาดพร้าว เขตบางกะปิ กรุงเทพฯ", 13.7656, 100.6430, "สายที่ 1"],
      ["ศูนย์การค้าแฟชั่นไอส์แลนด์", "ถนนรามอินทรา เขตคันนายาว กรุงเทพฯ", 13.8248, 100.6778, "สายที่ 2"],
      ["ลูกค้าสาขาอุดมสุข", "ถนนอุดมสุข แขวงบางนา เขตบางนา กรุงเทพฯ", 13.6781, 100.6139, "สายที่ 2"]
    ];

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Routing Template");

    try {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const btn = document.getElementById('btn-download-template');
      if (btn) {
        btn.href = url;
      }
    } catch (err) {
      console.error("ผิดพลาดในการสร้างไฟล์เทมเพลต: ", err);
    }
  }

  // ==========================================
  // 6. การตั้งค่าและสไลเดอร์ (Slider Config)
  // ==========================================
  function initSliderEvents() {
    const sliders = [
      { inputId: 'input-vehicles', labelId: 'val-vehicles', format: v => v },
      { inputId: 'input-stops', labelId: 'val-stops', format: v => v }
    ];

    sliders.forEach(slide => {
      const input = document.getElementById(slide.inputId);
      const label = document.getElementById(slide.labelId);
      if (input && label) {
        input.addEventListener('input', function() {
          label.textContent = slide.format(this.value);
        });
      }
    });
  }

  // ==========================================
  // 7. คำนวณและประมวลผลเส้นทาง (Solve Routing)
  // ==========================================
  function solveRoutingAndDisplay() {
    if (!appState.depot || appState.customers.length === 0) {
      showToast('กรุณานำเข้าข้อมูลก่อนคำนวณ', 'error');
      return;
    }

    const overlay = document.getElementById('loading-overlay');
    overlay.classList.add('active');

    setTimeout(() => {
      try {
        const config = {
          vehicleCount: document.getElementById('input-vehicles').value,
          maxStops: document.getElementById('input-stops').value,
          averageSpeed: document.getElementById('input-speed').value,
          serviceDuration: document.getElementById('input-service-time').value,
          isRoundTrip: document.getElementById('input-roundtrip').checked
        };

        const result = window.RoutingSolver.solveVRP(appState.depot, appState.customers, config);
        appState.solvedResult = result;

        overlay.classList.remove('active');

        renderRoutingResults(result);
        showToast('คำนวณจัดแยกสายส่งและเรียงลำดับเส้นทางเรียบร้อย', 'success');
        
        // บันทึกประวัติการจัดคิวสายวิ่งลงในฐานข้อมูลประวัติ
        saveRoutingToHistory(result);

        const resultTabBtn = document.querySelector('.tab-btn[data-tab="tab-results"]');
        if (resultTabBtn) resultTabBtn.click();

      } catch (err) {
        overlay.classList.remove('active');
        showToast('การคำนวณเกิดข้อผิดพลาด: ' + err.message, 'error');
        console.error(err);
      }
    }, 600);
  }

  // ==========================================
  // 8. แสดงผลลัพธ์ข้อมูลและสถิติ (Render Results UI)
  // ==========================================
  function renderRoutingResults(result) {
    document.getElementById('results-empty-state').style.display = 'none';
    document.getElementById('results-active-view').style.display = 'flex';

    document.getElementById('stat-total-distance').innerHTML = `${result.summary.totalDistance} <span class="stat-unit">กม.</span>`;
    
    const totalHours = Math.floor(result.summary.totalDuration / 60);
    const totalMinutes = result.summary.totalDuration % 60;
    const durationText = totalHours > 0 ? `${totalHours} ชม. ${totalMinutes} น.` : `${totalMinutes} นาที`;
    document.getElementById('stat-total-duration').innerHTML = `${durationText}`;
    
    const maxVehicles = document.getElementById('input-vehicles').value;
    document.getElementById('stat-vehicles-used').innerHTML = `${result.summary.vehiclesUsed} / ${maxVehicles} <span class="stat-unit">สาย</span>`;
    
    // อัปเดต stat-total-stops (จุดส่งสินค้ารวม) แทนที่ของเก่าน้ำหนัก
    document.getElementById('stat-total-stops').innerHTML = `${result.summary.totalCustomersCount} <span class="stat-unit">จุด</span>`;

    const unassignedWarning = document.getElementById('unassigned-warning');
    const unassignedList = document.getElementById('unassigned-list');
    const unassignedCount = document.getElementById('unassigned-count');

    if (result.unassigned && result.unassigned.length > 0) {
      unassignedWarning.style.display = 'block';
      unassignedCount.textContent = result.unassigned.length;
      unassignedList.innerHTML = '';
      
      result.unassigned.forEach(cust => {
        const tag = document.createElement('span');
        tag.className = 'unassigned-tag';
        tag.textContent = `${cust.name}`;
        unassignedList.appendChild(tag);
      });
    } else {
      unassignedWarning.style.display = 'none';
    }

    const accordion = document.getElementById('routes-accordion');
    accordion.innerHTML = '';

    result.routes.forEach((route, routeIndex) => {
      const routeHours = Math.floor(route.totalDuration / 60);
      const routeMins = Math.round(route.totalDuration % 60);
      const routeDurationStr = routeHours > 0 ? `${routeHours} ชม. ${routeMins} น.` : `${routeMins} น.`;

      const routeItem = document.createElement('div');
      routeItem.className = 'route-item';
      routeItem.dataset.routeIndex = routeIndex;
      
      routeItem.innerHTML = `
        <div class="route-header">
          <div class="route-title-section">
            <span class="route-color-indicator" style="background-color: ${route.color};"></span>
            <span class="route-title">${route.vehicleName} (${route.stops.length} จุดจอด)</span>
          </div>
          <div class="route-stats-summary">
            <span>🛣️ ${route.totalDistance.toFixed(1)} กม.</span>
            <span>⏱️ ${routeDurationStr}</span>
            <svg class="route-chevron" viewBox="0 0 24 24"><path d="M7.41 8.59L12 13.17l4.59-4.58L18 10l-6 6-6-6 1.41-1.41z"/></svg>
          </div>
        </div>
        
        <div class="route-body">
          <div class="stops-list">
            <div class="stop-node depot-node">
              <span class="stop-node-dot"></span>
              <div class="stop-node-info">
                <div class="stop-node-header">
                  <span class="stop-node-name">${appState.depot.name}</span>
                  <span class="stop-node-seq">จุดเริ่ม</span>
                </div>
                <span class="stop-node-detail">${appState.depot.address}</span>
              </div>
            </div>
            
            ${route.stops.map(stop => {
              const etaH = Math.floor(stop.etaOffset / 60);
              const etaM = Math.round(stop.etaOffset % 60);
              const etaString = etaH > 0 ? `+${etaH} ชม. ${etaM} น.` : `+${etaM} น.`;

              return `
                <div class="stop-node customer-node" data-cust-id="${stop.customer.id}">
                  <span class="stop-node-dot" style="background-color: ${route.color}; box-shadow: 0 0 6px ${route.color};"></span>
                  <div class="stop-node-info">
                    <div class="stop-node-header">
                      <span class="stop-node-name">${stop.customer.name}</span>
                      <span class="stop-node-seq">${stop.sequence}</span>
                    </div>
                    <span class="stop-node-detail">${stop.customer.address}</span>
                    <div class="stop-node-meta">
                      <span>📏 ระยะจอด: ${stop.legDistance.toFixed(1)} กม.</span>
                      <span class="stop-node-eta">🕒 ETA: <b>${etaString}</b></span>
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
            
            ${document.getElementById('input-roundtrip').checked ? `
              <div class="stop-node depot-node">
                <span class="stop-node-dot"></span>
                <div class="stop-node-info">
                  <div class="stop-node-header">
                    <span class="stop-node-name">${appState.depot.name}</span>
                    <span class="stop-node-seq">จุดกลับ</span>
                  </div>
                  <span class="stop-node-detail">${appState.depot.address}</span>
                  <div class="stop-node-meta">
                    <span>📏 ระยะขากลับ: ${(route.totalDistance - route.stops[route.stops.length - 1].cumulativeDistance).toFixed(1)} กม.</span>
                  </div>
                </div>
              </div>
            ` : ''}
          </div>
        </div>
      `;

      const header = routeItem.querySelector('.route-header');
      header.addEventListener('click', function() {
        const activeItems = accordion.querySelectorAll('.route-item.active');
        activeItems.forEach(item => {
          if (item !== routeItem) item.classList.remove('active');
        });
        
        routeItem.classList.toggle('active');
        highlightRouteOnMap(route.id);
      });

      const custNodes = routeItem.querySelectorAll('.customer-node');
      custNodes.forEach(node => {
        const custId = parseInt(node.getAttribute('data-cust-id'));
        
        node.addEventListener('mouseenter', () => {
          const marker = appState.mapMarkers[custId];
          if (marker) {
            marker.openPopup();
            appState.map.panTo(marker.getLatLng());
          }
        });
      });

      accordion.appendChild(routeItem);
// Initialise SortableJS for stops within this route
const stopsList = routeItem.querySelector('.stops-list');
if (window.Sortable && stopsList) {
  new Sortable(stopsList, {
    animation: 150,
    ghostClass: 'sortable-ghost',
    draggable: '.stop-node',
    onEnd: (evt) => {
      const oldIdx = evt.oldIndex;
      const newIdx = evt.newIndex;
      if (oldIdx === undefined || newIdx === undefined) return;
      // Adjust for depot node at index 0
      const stopOldIdx = oldIdx - 1;
      const stopNewIdx = newIdx - 1;
      if (stopOldIdx < 0 || stopNewIdx < 0) return; // ignore depot moves
      const routeIdx = parseInt(routeItem.dataset.routeIndex);
      const routeData = appState.solvedResult.routes[routeIdx];
      const moved = routeData.stops.splice(stopOldIdx, 1)[0];
      routeData.stops.splice(stopNewIdx, 0, moved);
      // Update sequence numbers
      routeData.stops.forEach((s, i) => s.sequence = i + 1);
      // Re‑render UI and persist changes
      renderRoutingResults(appState.solvedResult);
      // Keep the moved route expanded after re‑render
      const accordion = document.getElementById('routes-accordion');
      const reopened = accordion.querySelector(`.route-item[data-route-index="${routeIdx}"]`);
      if (reopened) reopened.classList.add('active');
      db.routingHistory.put({ timestamp: Date.now(), result: appState.solvedResult });
    }
  });
}
    });

    // Initialise SortableJS for route reordering
    if (window.Sortable) {
      new Sortable(accordion, {
        animation: 150,
        ghostClass: 'sortable-ghost',
        onEnd: (evt) => {
          const oldIdx = evt.oldIndex;
          const newIdx = evt.newIndex;
          if (oldIdx === undefined || newIdx === undefined) return;
          const moved = appState.solvedResult.routes.splice(oldIdx, 1)[0];
          appState.solvedResult.routes.splice(newIdx, 0, moved);
          renderRoutingResults(appState.solvedResult);
          db.routingHistory.put({ timestamp: Date.now(), result: appState.solvedResult });
        }
      });
    }

    plotSolvedRoutesOnMap(result);
  }

  // ==========================================
  // 9. พล็อตเส้นทางตามโครงข่ายถนนจริง (OSRM Integration)
  // ==========================================
  function plotSolvedRoutesOnMap(result) {
    appState.markersLayer.clearLayers();
    appState.routesLayer.clearLayers();
    appState.mapMarkers = {};

    // พล็อตคลังสินค้าเริ่มต้น
    const depotIcon = L.divIcon({
      className: 'custom-div-icon',
      html: `<div class="depot-marker-pin">
              <svg viewBox="0 0 24 24"><path d="M19 4H5c-1.11 0-2 .9-2 2v12c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm-9 14H5v-5h5v5zm0-6H5V7h5v5zm9 6h-5v-5h5v5zm0-6h-5V7h5v5z"/></svg>
            </div>`,
      iconSize: [28, 28],
      iconAnchor: [14, 14]
    });

    const depotMarker = L.marker([appState.depot.lat, appState.depot.lng], { icon: depotIcon })
      .bindPopup(`<div class="popup-title">${appState.depot.name}</div><div class="popup-detail">คลังสินค้าศูนย์กลาง (Depot)</div>`)
      .addTo(appState.markersLayer);
    
    appState.mapMarkers[appState.depot.id] = depotMarker;

    // พล็อตแต่ละเส้นทางเดินรถ
    result.routes.forEach(route => {
      const straightPoints = [];
      straightPoints.push([appState.depot.lat, appState.depot.lng]);
      
      route.stops.forEach(stop => {
        straightPoints.push([stop.customer.lat, stop.customer.lng]);
      });

      if (document.getElementById('input-roundtrip').checked) {
        straightPoints.push([appState.depot.lat, appState.depot.lng]);
      }

      // วาดเส้นตรง Polyline ชั่วคราว (Fallback) ไว้ก่อนเพื่อให้เห็นผลทันที
      const routeLine = L.polyline(straightPoints, {
        color: route.color,
        weight: 4,
        opacity: 0.65,
        className: `route-polyline-${route.id}`
      }).addTo(appState.routesLayer);

      routeLine.bindPopup(`<b>${route.vehicleName}</b><br>ระยะวิ่งรวม: ${route.totalDistance.toFixed(1)} กม.<br>จำนวนส่ง: ${route.stops.length} จุดจอด`);

      // เรียกดึงข้อมูล OSRM เพื่อเปลี่ยนเป็นเส้นวิ่งตามถนนจริง
      fetchOSRMRoute(straightPoints, routeLine);

      // วาดหมุดเรียงลำดับจุดจอดสำหรับสายส่งนี้
      route.stops.forEach(stop => {
        const markerIcon = L.divIcon({
          className: 'custom-div-icon',
          html: `
            <div class="map-marker-pin" style="background-color: ${route.color};"></div>
            <div class="map-marker-label">${stop.sequence}</div>
          `,
          iconSize: [24, 24],
          iconAnchor: [12, 12]
        });

        const etaH = Math.floor(stop.etaOffset / 60);
        const etaM = Math.round(stop.etaOffset % 60);
        const etaString = etaH > 0 ? `${etaH} ชม. ${etaM} น.` : `${etaM} น.`;

        const marker = L.marker([stop.customer.lat, stop.customer.lng], { icon: markerIcon })
          .bindPopup(`
            <div class="popup-title">${stop.customer.name}</div>
            <div class="popup-detail">${stop.customer.address}</div>
            <div class="popup-detail"><b>ขบวนเดินรถ:</b> ${route.vehicleName}</div>
            <div class="popup-detail"><b>ลำดับส่งของ:</b> จุดที่ ${stop.sequence}</div>
            <div class="popup-detail"><b>ระยะสะสมจากคลัง:</b> ${stop.cumulativeDistance.toFixed(1)} กม.</div>
            <div class="popup-tag" style="background-color: ${route.color}25; color: ${route.color}; border: 1px solid ${route.color}50;">🕒 ETA จากคลัง: +${etaString}</div>
          `)
          .addTo(appState.markersLayer);

        appState.mapMarkers[stop.customer.id] = marker;
      });
    });

    // แสดงลูกค้าที่ไม่สามารถวิ่งส่งได้
    if (result.unassigned && result.unassigned.length > 0) {
      result.unassigned.forEach(cust => {
        const markerIcon = L.divIcon({
          className: 'custom-div-icon',
          html: `<div class="map-marker-pin" style="background-color: #ef4444; box-shadow: 0 0 10px #ef4444;"></div><div class="map-marker-label">⚠️</div>`,
          iconSize: [24, 24],
          iconAnchor: [12, 12]
        });

        const marker = L.marker([cust.lat, cust.lng], { icon: markerIcon })
          .bindPopup(`
            <div class="popup-title" style="color: #ef4444;">⚠️ จุดตกหล่นนอกแผนการเดินรถ</div>
            <div class="popup-detail">${cust.address}</div>
            <div class="popup-detail" style="color:#f87171; font-weight:500; margin-top: 5px;">*เนื่องจากจำนวนสายรถมีไม่พอกลุ่มพิกัด หรือเกินโควต้าจุดจอดสูงสุดต่อคัน</div>
          `)
          .addTo(appState.markersLayer);

        appState.mapMarkers[cust.id] = marker;
      });
    }

    if (appState.markersLayer.getLayers().length > 0) {
      appState.map.fitBounds(appState.markersLayer.getBounds(), { padding: [50, 50] });
    }
  }

  // เรียกดึงพิกัดถนนและอัปเดตเส้นเดินทางจาก OSRM API (ขับขี่รถยนต์)
  async function fetchOSRMRoute(coordinates, polylineObject) {
    if (coordinates.length < 2) return;
    
    // แปลงพิกัดเป็นฟอร์แมต OSRM (lng,lat คั่นด้วยเซมิโคลอน)
    const coordsString = coordinates.map(coord => `${coord[1]},${coord[0]}`).join(';');
    const osrmUrl = `https://router.project-osrm.org/route/v1/driving/${coordsString}?overview=full&geometries=geojson`;

    try {
      const response = await fetch(osrmUrl);
      if (!response.ok) {
        throw new Error('OSRM API response status error');
      }
      
      const data = await response.json();
      if (data.code === 'Ok' && data.routes && data.routes.length > 0) {
        // แปลงพิกัดกลับเป็น Leaflet [lat, lng]
        const roadPoints = data.routes[0].geometry.coordinates.map(coord => [coord[1], coord[0]]);
        
        // อัปเดตพิกัดเส้นโยงบนแผนที่ในส่วนของ Polyline ที่แสดงอยู่ทันที
        polylineObject.setLatLngs(roadPoints);
      }
    } catch (err) {
      console.warn('ล้มเหลวในการเชื่อมต่อระบบ OSRM หรือเกิดข้อผิดพลาด ระบบถอยกลับไปวาดเส้นตรงพิกัดตามปกติ:', err.message);
    }
  }

  function highlightRouteOnMap(activeRouteId) {
    appState.solvedResult.routes.forEach(route => {
      const lines = document.querySelectorAll(`.route-polyline-${route.id}`);
      lines.forEach(path => {
        if (route.id === activeRouteId) {
          path.setAttribute('stroke-width', '7');
          path.setAttribute('stroke-opacity', '1');
        } else {
          path.setAttribute('stroke-width', '3');
          path.setAttribute('stroke-opacity', '0.2');
        }
      });
    });
  }

  // ==========================================
  // 10. ส่งออกข้อมูลผลลัพธ์เป็น Excel (Export)
  // ==========================================
  function exportResultsToExcel() {
    if (!appState.solvedResult) return;

    const workbook = XLSX.utils.book_new();

    // 1. สร้างตารางแผนรวม (Tab 1: Dashboard)
    const summaryData = [
      ["หัวข้อรายงานสรุปแผนเดินรถแบบจัดสาย", "ข้อมูลสถิติ"],
      ["ระยะทางทั้งหมดที่วิ่งรวม (กิโลเมตร)", appState.solvedResult.summary.totalDistance],
      ["ประมาณการเวลาเดินทางรวม (ชั่วโมง)", (appState.solvedResult.summary.totalDuration / 60).toFixed(1)],
      ["จำนวนสายส่งที่จัดรถวิ่ง (สาย/คัน)", appState.solvedResult.summary.vehiclesUsed],
      ["จำนวนจุดส่งมอบที่นำเข้ารวม (จุด)", appState.solvedResult.summary.totalCustomersCount],
      ["จำนวนจุดส่งที่วิ่งสำเร็จในคิว (จุด)", appState.solvedResult.summary.totalCustomersCount - appState.solvedResult.unassigned.length],
      ["จำนวนจุดส่งสินค้าที่หล่นคิว/ค้างจัด", appState.solvedResult.unassigned.length]
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(workbook, wsSummary, "สรุปภาพรวมแผน");

    // 2. สร้างรายละเอียดลำดับการขนส่งแต่ละคัน (Tab 2: Details)
    const detailHeaders = [
      "สายส่งสินค้า", 
      "คิวจอด", 
      "ชื่อสถานที่ / ลูกค้า", 
      "รายละเอียดที่อยู่สถานที่", 
      "ระยะทางจากจุดก่อนหน้า (กม.)", 
      "ระยะทางสะสมรวม (กม.)", 
      "เวลาประมาณการที่จะถึง (นาทีนับจากจุดเริ่ม)"
    ];

    const detailRows = [detailHeaders];

    appState.solvedResult.routes.forEach(route => {
      detailRows.push([
        route.vehicleName,
        "คลังเริ่มต้น",
        appState.depot.name,
        appState.depot.address,
        0,
        0,
        0
      ]);

      route.stops.forEach(stop => {
        detailRows.push([
          route.vehicleName,
          stop.sequence,
          stop.customer.name,
          stop.customer.address,
          parseFloat(stop.legDistance.toFixed(2)),
          parseFloat(stop.cumulativeDistance.toFixed(2)),
          Math.round(stop.etaOffset)
        ]);
      });

      if (document.getElementById('input-roundtrip').checked) {
        const lastStop = route.stops[route.stops.length - 1];
        const returnDist = route.totalDistance - lastStop.cumulativeDistance;
        const totalDuration = route.totalDuration;
        
        detailRows.push([
          route.vehicleName,
          "คลังขากลับ",
          appState.depot.name,
          appState.depot.address,
          parseFloat(returnDist.toFixed(2)),
          parseFloat(route.totalDistance.toFixed(2)),
          Math.round(totalDuration)
        ]);
      }
      
      detailRows.push([]);
    });

    if (appState.solvedResult.unassigned.length > 0) {
      detailRows.push(["--- จุดจัดส่งค้างจัดมอบเนื่องจากเกินพิกัดจำนวนรถ/โควต้าจุดจอดสูงสุด ---"]);
      appState.solvedResult.unassigned.forEach(cust => {
        detailRows.push([
          "ค้างส่ง (Unassigned)",
          "-",
          cust.name,
          cust.address,
          0,
          0,
          0
        ]);
      });
    }

    const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
    XLSX.utils.book_append_sheet(workbook, wsDetail, "รายละเอียดคิวส่งของ");

    XLSX.writeFile(workbook, `แผนจัดลำดับสายส่งสินค้า_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast('ดาวน์โหลดรายงานแผนเดินรถสำเร็จ', 'success');
  }

  // ==========================================
  // 11. ฟังก์ชันตัวช่วยแจ้งเตือน (Notifications Toast)
  // ==========================================
  function showToast(message, type = 'success') {
    const toast = document.getElementById('notification');
    const toastText = document.getElementById('notification-text');
    const toastIcon = document.getElementById('notification-icon');

    toastText.textContent = message;
    toast.className = 'notification';
    
    if (type === 'error') {
      toast.classList.add('error');
      toastIcon.innerHTML = `<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>`;
    } else if (type === 'info') {
      toast.classList.add('info');
      toastIcon.innerHTML = `<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/>`;
    } else {
      toastIcon.innerHTML = `<path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>`;
    }

    toast.classList.add('show');

    setTimeout(() => {
      toast.classList.remove('show');
    }, 3500);
  }

  // ==========================================
  // 11.5. ระบบฐานข้อมูล IndexedDB (Dexie.js)
  // ==========================================
  let db;

  function initDatabase() {
    try {
      db = new Dexie("SmartRouteDB");
      db.version(1).stores({
        configStore: "key", // key: value สำหรับเก็บค่า config ล่าสุด
        savedLocations: "++id, name, date, locationsCount", // เก็บบล็อกข้อมูลจุดส่งลูกค้า
        routingHistory: "++id, name, date, distance, duration, stopsCount, vehiclesCount" // เก็บคิวประวัติการจัดคิววิ่งรถ
      });
      
      // โหลดการตั้งค่าเก่าจาก DB และอัปเดต UI ประวัติต่างๆ
      loadSavedConfig();
      loadSavedSetsList();
      loadHistoryList();

      // เพิ่ม Event Handlers สำหรับปุ่มต่างๆ ในแถบ DB
      document.getElementById('btn-save-current-to-db').addEventListener('click', saveCurrentLocationsToDB);
      document.getElementById('btn-clear-all-db').addEventListener('click', clearAllDatabase);

    } catch (err) {
      console.warn("ไม่สามารถรัน IndexedDB ได้: ", err.message);
    }
  }

  // โหลดค่า config เก่าจาก DB
  async function loadSavedConfig() {
    try {
      const keys = ['vehicles', 'stops', 'speed', 'service-time', 'roundtrip', 'depot-source', 'custom-depot-name', 'custom-depot-lat', 'custom-depot-lng'];
      
      for (let key of keys) {
        const item = await db.configStore.get(key);
        if (item) {
          const val = item.value;
          if (key === 'vehicles') {
            document.getElementById('input-vehicles').value = val;
            document.getElementById('val-vehicles').textContent = val;
          } else if (key === 'stops') {
            document.getElementById('input-stops').value = val;
            document.getElementById('val-stops').textContent = val;
          } else if (key === 'speed') {
            document.getElementById('input-speed').value = val;
          } else if (key === 'service-time') {
            document.getElementById('input-service-time').value = val;
          } else if (key === 'roundtrip') {
            document.getElementById('input-roundtrip').checked = (val === 'true');
          } else if (key === 'depot-source') {
            if (val === 'custom') {
              document.getElementById('depot-source-custom').click();
            } else {
              document.getElementById('depot-source-import').click();
            }
          } else if (key === 'custom-depot-name') {
            document.getElementById('custom-depot-name').value = val;
          } else if (key === 'custom-depot-lat') {
            document.getElementById('custom-depot-lat').value = val;
          } else if (key === 'custom-depot-lng') {
            document.getElementById('custom-depot-lng').value = val;
          }
        }
      }
      
      // บังคับอัปเดต state Depot จาก config โหลดใหม่
      updateDepotAndCustomers();
    } catch (err) {
      console.error("ผิดพลาดในการโหลด Config: ", err);
    }
  }

  // บันทึก Config ลง DB
  function saveConfig(key, value) {
    if (!db) return;
    db.configStore.put({ key: key, value: String(value) }).catch(err => {
      console.error("ล้มเหลวในการเซฟ config: ", err);
    });
  }

  // ผูกตัวบันทึก Config อัตโนมัติในสไลเดอร์และช่องป้อนข้อมูล
  document.getElementById('input-vehicles').addEventListener('change', function() { saveConfig('vehicles', this.value); });
  document.getElementById('input-stops').addEventListener('change', function() { saveConfig('stops', this.value); });
  document.getElementById('input-speed').addEventListener('change', function() { saveConfig('speed', this.value); });
  document.getElementById('input-service-time').addEventListener('change', function() { saveConfig('service-time', this.value); });
  document.getElementById('input-roundtrip').addEventListener('change', function() { saveConfig('roundtrip', this.checked); });
  document.getElementById('custom-depot-name').addEventListener('change', function() { saveConfig('custom-depot-name', this.value); });
  document.getElementById('custom-depot-lat').addEventListener('change', function() { saveConfig('custom-depot-lat', this.value); });
  document.getElementById('custom-depot-lng').addEventListener('change', function() { saveConfig('custom-depot-lng', this.value); });
  
  if (document.getElementById('depot-source-import')) {
    document.getElementById('depot-source-import').addEventListener('click', () => saveConfig('depot-source', 'import'));
    document.getElementById('depot-source-custom').addEventListener('click', () => saveConfig('depot-source', 'custom'));
  }

  // โหลดรายการเซฟชุดพิกัด (Locations Sets)
  async function loadSavedSetsList() {
    if (!db) return;
    const listContainer = document.getElementById('saved-sets-list');
    listContainer.innerHTML = '';
    
    try {
      const sets = await db.savedLocations.toArray();
      if (sets.length === 0) {
        listContainer.innerHTML = '<div class="empty-state" style="padding: 10px; font-size: 11px;">ไม่มีชุดพิกัดที่บันทึกไว้</div>';
        return;
      }
      
      sets.forEach(setItem => {
        const div = document.createElement('div');
        div.style.cssText = 'padding: 8px; display: flex; flex-direction: column; gap: 4px; background: rgba(255,255,255,0.015); border: 1px solid var(--border-glass); border-radius: 6px;';
        
        div.innerHTML = `
          <div style="display: flex; justify-content: space-between; align-items: baseline; gap: 10px;">
            <span style="font-size: 12px; font-weight: 600; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 170px;" title="${setItem.name}">${setItem.name}</span>
            <span style="font-size: 9px; color: var(--text-muted);">${setItem.date.split(' ')[0]}</span>
          </div>
          <div style="font-size: 10px; color: var(--text-sub);">📍 มีจุดจอด ${setItem.locationsCount} จุด</div>
          <div style="display: flex; gap: 6px; margin-top: 4px;">
            <button class="btn btn-sm" style="flex: 1; padding: 4px;" id="btn-load-set-${setItem.id}">โหลดข้อมูล</button>
            <button class="btn btn-sm btn-remove-file" style="padding: 4px; color: var(--color-danger);" id="btn-delete-set-${setItem.id}">ลบ</button>
          </div>
        `;
        listContainer.appendChild(div);
        
        // ผูกคลิกโหลด
        document.getElementById(`btn-load-set-${setItem.id}`).addEventListener('click', () => {
          loadLocationsSetFromDB(setItem.id);
        });
        // ผูกคลิกลบ
        document.getElementById(`btn-delete-set-${setItem.id}`).addEventListener('click', (e) => {
          e.stopPropagation();
          deleteLocationsSetFromDB(setItem.id);
        });
      });
    } catch (err) {
      console.error(err);
    }
  }

  // โหลดรายการประวัติจัดสายรถ (Routing History)
  async function loadHistoryList() {
    if (!db) return;
    const listContainer = document.getElementById('routing-history-list');
    listContainer.innerHTML = '';
    
    try {
      // เรียงจากประวัติใหม่สุดลงไปเก่าสุด
      const historyItems = await db.routingHistory.reverse().toArray();
      if (historyItems.length === 0) {
        listContainer.innerHTML = '<div class="empty-state" style="padding: 10px; font-size: 11px;">ไม่มีประวัติการคำนวณ</div>';
        return;
      }

      historyItems.forEach(item => {
        const div = document.createElement('div');
        div.style.cssText = 'padding: 8px; display: flex; flex-direction: column; gap: 4px; background: rgba(255,255,255,0.015); border: 1px solid var(--border-glass); border-radius: 6px;';
        
        div.innerHTML = `
          <div style="display: flex; justify-content: space-between; align-items: baseline; gap: 10px;">
            <span style="font-size: 12px; font-weight: 600; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 170px;" title="${item.name}">${item.name}</span>
            <span style="font-size: 9px; color: var(--text-muted);">${item.date}</span>
          </div>
          <div style="font-size: 10px; color: var(--text-sub);">🛣️ ${item.distance} กม. | 📍 ${item.stopsCount} จุด | 🚗 ${item.vehiclesCount} สาย</div>
          <div style="display: flex; gap: 6px; margin-top: 4px;">
            <button class="btn btn-sm" style="flex: 1; padding: 4px;" id="btn-load-hist-${item.id}">เรียกดูแผนที่</button>
            <button class="btn btn-sm btn-remove-file" style="padding: 4px; color: var(--color-danger);" id="btn-delete-hist-${item.id}">ลบ</button>
          </div>
        `;
        listContainer.appendChild(div);

        document.getElementById(`btn-load-hist-${item.id}`).addEventListener('click', () => {
          loadHistoryItemFromDB(item.id);
        });

        document.getElementById(`btn-delete-hist-${item.id}`).addEventListener('click', (e) => {
          e.stopPropagation();
          deleteHistoryItemFromDB(item.id);
        });
      });
    } catch (err) {
      console.error(err);
    }
  }

  // บันทึกกลุ่มจุดส่งปัจจุบันลง DB
  async function saveCurrentLocationsToDB() {
    if (!appState.importedData || appState.importedData.length === 0) {
      showToast('กรุณานำเข้าหรือโหลดชุดข้อมูลก่อนเซฟ', 'error');
      return;
    }

    const nameInput = document.getElementById('input-save-db-name').value.trim();
    const name = nameInput || `กลุ่มข้อมูล_${new Date().toLocaleDateString('th-TH')} (${appState.importedData.length} จุด)`;

    try {
      const setItem = {
        name: name,
        date: new Date().toLocaleString('th-TH'),
        locationsCount: appState.importedData.length,
        importedData: JSON.parse(JSON.stringify(appState.importedData))
      };

      await db.savedLocations.add(setItem);
      document.getElementById('input-save-db-name').value = '';
      loadSavedSetsList();
      showToast('บันทึกชุดจุดจอดส่งสินค้าลง Database เรียบร้อย', 'success');
    } catch (err) {
      showToast('ผิดพลาดในการเซฟข้อมูล: ' + err.message, 'error');
    }
  }

  // ออโต้เซฟผลคำนวณลงประวัติ
  async function saveRoutingToHistory(result) {
    if (!db || !result) return;
    
    try {
      const name = appState.fileName ? appState.fileName.replace('.xlsx', '').replace('.xls', '').replace('.csv', '') : 'จัดสายส่งชั่วคราว';
      const historyItem = {
        name: name,
        date: new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }) + " " + new Date().toLocaleDateString('th-TH'),
        distance: result.summary.totalDistance,
        duration: result.summary.totalDuration,
        stopsCount: result.summary.totalCustomersCount,
        vehiclesCount: result.summary.vehiclesUsed,
        depot: JSON.parse(JSON.stringify(appState.depot)),
        routes: JSON.parse(JSON.stringify(result.routes)),
        unassigned: JSON.parse(JSON.stringify(result.unassigned))
      };

      await db.routingHistory.add(historyItem);
      loadHistoryList();
    } catch (err) {
      console.warn("ไม่สามารถบันทึกประวัติการเดินรถได้: ", err.message);
    }
  }

  // โหลดพิกัดจุดส่งจากเซฟใน DB
  async function loadLocationsSetFromDB(id) {
    try {
      const setItem = await db.savedLocations.get(id);
      if (!setItem) {
        showToast('ไม่พบข้อมูลชุดนี้ในระบบ', 'error');
        return;
      }

      appState.importedData = JSON.parse(JSON.stringify(setItem.importedData));
      appState.fileName = setItem.name;

      updateDataUIs();
      showToast(`โหลดชุดพิกัด "${setItem.name}" สำเร็จ`, 'success');

      // สลับไปแท็บ 1
      const importTabBtn = document.querySelector('.tab-btn[data-tab="tab-import"]');
      if (importTabBtn) importTabBtn.click();
    } catch (err) {
      showToast('ผิดพลาดในการโหลดข้อมูล: ' + err.message, 'error');
    }
  }

  // โหลดผลลัพธ์ประวัติการเดินรถจาก DB มาพล็อต
  async function loadHistoryItemFromDB(id) {
    try {
      const hist = await db.routingHistory.get(id);
      if (!hist) {
        showToast('ไม่พบประวัติการคำนวณรายการนี้', 'error');
        return;
      }

      // โหลดพิกัด
      appState.depot = JSON.parse(JSON.stringify(hist.depot));
      appState.fileName = `ประวัติ: ${hist.name}`;
      
      // สร้างผลลัพธ์จัดคิววิ่งรถ
      const reconstructedResult = {
        routes: JSON.parse(JSON.stringify(hist.routes)),
        unassigned: JSON.parse(JSON.stringify(hist.unassigned)),
        summary: {
          totalDistance: hist.distance,
          totalDuration: hist.duration,
          vehiclesUsed: hist.vehiclesCount,
          totalCustomersCount: hist.stopsCount
        }
      };

      appState.solvedResult = reconstructedResult;

      // สลับแท็บและพล็อต
      renderRoutingResults(reconstructedResult);
      showToast(`เรียกดูประวัติแผนวิ่งรถ "${hist.name}" สำเร็จ`, 'success');

      // สลับไปแท็บผลลัพธ์
      const resultTabBtn = document.querySelector('.tab-btn[data-tab="tab-results"]');
      if (resultTabBtn) resultTabBtn.click();
    } catch (err) {
      showToast('ผิดพลาดในการโหลดประวัติ: ' + err.message, 'error');
    }
  }

  // ลบชุดพิกัด
  async function deleteLocationsSetFromDB(id) {
    try {
      await db.savedLocations.delete(id);
      loadSavedSetsList();
      showToast('ลบชุดพิกัดจุดส่งสินค้าแล้ว', 'info');
    } catch (err) {
      showToast('ผิดพลาดในการลบ: ' + err.message, 'error');
    }
  }

  // ลบประวัติเดินรถ
  async function deleteHistoryItemFromDB(id) {
    try {
      await db.routingHistory.delete(id);
      loadHistoryList();
      showToast('ลบประวัติการเดินรถรายการนี้แล้ว', 'info');
    } catch (err) {
      showToast('ผิดพลาดในการลบประวัติ: ' + err.message, 'error');
    }
  }

  // ล้างฐานข้อมูลประวัติและข้อมูลทั้งหมด
  async function clearAllDatabase() {
    if (confirm('⚠️ คำเตือน: คุณต้องการล้างฐานข้อมูลประวัติ ค่าติดตั้ง และชุดจุดส่งสินค้าทั้งหมดในเครื่องใช่หรือไม่?')) {
      try {
        await Promise.all([
          db.configStore.clear(),
          db.savedLocations.clear(),
          db.routingHistory.clear()
        ]);
        
        showToast('ล้างฐานข้อมูลสำเร็จแล้ว ระบบกำลังรีเซ็ตแอป', 'success');
        
        setTimeout(() => {
          window.location.reload();
        }, 1500);
      } catch (err) {
        showToast('ผิดพลาดในการล้างระบบ: ' + err.message, 'error');
      }
    }
  }

  // ==========================================
  // 12. เรียกติดตั้งแอป (Self-Init)
  // ==========================================
  initApp();

});
