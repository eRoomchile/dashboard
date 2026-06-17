// ========================
// 전역 변수
// ========================
var NOMBRE_INVENTARIO = "Inventario";
var NOMBRE_VENTAS = "Ventas";
var NOMBRE_INGRESOS = "Ingresos";
var NOMBRE_REGISTRO = "Registro de ventas";
var NOMBRE_CLIENTES = "List de cliente";
var ID_통합재고 = "1TUxCrEVvgDmiutoUkKfqpF7zUgh23ajFhm3PMQY6QMM";

var _clientCache = null;
var _invCache = null;

function getInvData() {
  if (_invCache) return _invCache;
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('invData');
    if (cached) { _invCache = JSON.parse(cached); return _invCache; }
  } catch(e) {}
  var invSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) return null;
  _invCache = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 7).getValues();
  try { CacheService.getScriptCache().put('invData', JSON.stringify(_invCache), 300); } catch(e) {}
  return _invCache;
}

function invCacheClear() {
  _invCache = null;
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('invData');
    cache.remove('inventarioUIData'); // ← UI 캐시도 초기화
  } catch(e) {}
}

function getMayor(cliente) {
  if (!_clientCache) {
    var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_CLIENTES);
    if (!clientSheet) return false;
    _clientCache = clientSheet.getDataRange().getValues();
  }
  for (var c = 0; c < _clientCache.length; c++) {
    if (String(_clientCache[c][0]).toLowerCase().trim() === cliente &&
        String(_clientCache[c][1]).toLowerCase().trim() === "mayor") return true;
  }
  return false;
}

// ========================
// onEdit
// ========================
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var range = e.range;
    var sheet = range.getSheet();
    var row = range.getRow();
    var col = range.getColumn();
    var sheetName = sheet.getName();
    if (sheetName !== NOMBRE_VENTAS &&
        sheetName !== NOMBRE_INGRESOS &&
        sheetName !== NOMBRE_INVENTARIO &&
        sheetName !== "재고수동조정" &&
        sheetName !== "재고실사") return;
    if (row < 2) return;
    _clientCache = null;
    if (sheetName === NOMBRE_VENTAS) {
      if (col === 4) procesarBarcodeVentas(e);
      if (col === 2) procesarCambioCliente(e);
      if (col === 7) procesarCantidad(e);
      if (col === 8) procesarPrecioManual(e);
    }
    if (sheetName === NOMBRE_INGRESOS) {
      if (col === 2) procesarBarcodeIngresos(e);
      if (col === 5 || col === 6) {
        var qty = sheet.getRange(row, 5).getValue();
        var costo = sheet.getRange(row, 6).getValue();
        if (qty && costo) sheet.getRange(row, 7).setValue(qty * costo).setNumberFormat("#,##0");
      }
    }
    if (sheetName === NOMBRE_INVENTARIO && row >= 2) {
      invCacheClear();
    }
    if (sheetName === "재고수동조정" && col === 2) {
      var barcode = String(range.getValue()).replace(/[^0-9]/g, '');
      if (barcode) {
        range.setValue(barcode);
        var invData = getInvData();
        if (invData) {
          for (var i = 0; i < invData.length; i++) {
            if (String(invData[i][0]).replace(/[^0-9]/g, '') === barcode) {
              sheet.getRange(row, 1).setValue(new Date()).setNumberFormat("yyyy-MM-dd HH:mm");
              sheet.getRange(row, 3).setValue(invData[i][1]);
              sheet.getRange(row, 4).setValue(invData[i][2]);
              sheet.getRange(row, 5).setValue(invData[i][6]);
              sheet.setActiveRange(sheet.getRange(row, 6));
              break;
            }
          }
        }
      }
    }
    if (sheetName === "재고실사") 재고실사onEdit(e);
    행색상자동적용(sheet, row);
  } catch (error) {
    console.log("onEdit 오류: " + error.toString());
  }
}

// ========================
// onOpen 메뉴
// ========================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏪 NAOS Ventas")
    .addItem("🔍 Buscar producto / 상품 검색", "상품검색사이드바")
    .addSeparator()
    .addItem("🔍 Verificar código (Venta) / 바코드 확인 (판매)", "선택셀바코드확인")
    .addItem("📦 Verificar código (Ingreso) / 바코드 확인 (입고)", "입고바코드확인")
    .addItem("📦 Procesar ingreso / 입고완료 처리", "procesarIngresos")
    .addSeparator()
    .addItem("🧾 Imprimir BOLETA / BOLETA 출력", "출력BOLETA")
    .addItem("🧾 Imprimir FACTURA / FACTURA 출력", "출력FACTURA")
    .addSeparator()
    .addItem("📊 Reporte cierre / 일일마감 보고서", "일일마감보고서생성")
    .addSeparator()
    .addItem("🏭 거래처관리로 보내기", "inventarioToGeorae") 
    .addToUi();

  ui.createMenu("⚙️ NAOS Gestión")
    .addItem("⚠️ Verificar stock bajo / 부족재고 확인", "부족재고표시")
    .addSeparator()
    .addItem("📋 Ajuste manual stock / 재고수동조정", "재고수동조정열기")
    .addItem("✅ Aplicar ajuste / 재고조정 반영", "재고수동조정처리")
    .addSeparator()
    .addItem("📦 Iniciar inventario / 재고실사 시작", "재고실사시트생성")
    .addItem("🔍 Ver diferencias / 차이 항목만 보기", "재고실사차이표시")
    .addItem("📋 Ver todo / 재고실사 전체보기", "재고실사전체보기")
    .addItem("✅ Aplicar inventario / 재고실사 결과반영", "재고실사결과반영")
    .addSeparator()
    .addItem("📊 Análisis ventas / 매출분석 업데이트", "매출분석업데이트")
    .addSeparator()
    .addItem("🔢 Generar código / 바코드 자동생성", "바코드직접생성")
    .addItem("🎨 Aplicar diseño / 시트 디자인 적용", "시트디자인적용")
    .addItem("📷 Conectar fotos / 사진 연결", "사진열추가_Inventario")
    .addItem("📷 Mostrar/Ocultar fotos / 사진열 숨기기/보이기", "사진열토글")
    .addItem("🗑️ Limpiar columna H / H열 초기화", "H열초기화")
    .addSeparator()
    .addItem("🙈 Ocultar hojas / 시트 숨기기", "시트숨기기관리")
    .addItem("👁️ Mostrar hojas / 시트 보이기", "시트보이기관리")
    .addItem("🔒 시트 권한 설정", "시트권한일괄설정")
    .addItem("📊 ML비교결과 업데이트", "updateMLComparison")
    .addToUi();
    ui.createMenu("📦 발주서 연동")
    .addItem("📤 Eroom 발주서 → NAOS 판매 전송", "sendEroomOrderToNAOS")
    .addToUi();
    ui.createMenu("🛒 ML 판매자동화")
    .addItem("📥 ML 판매 가져오기 + 재고차감 (API)", "fetchMLSalesAndDeduct")
    .addItem("🔄 토큰 수동 갱신", "manualRefreshMLToken")
    .addSeparator()
    .addItem("✏️ 수동입력 재고 차감", "mlSalesManualDeduct")
    .addItem("📋 묶음설정 시트 열기", "묶음설정시트열기")
    .addSeparator()
    .addItem("🔑 토큰 설정 확인", "mlTokenSetupHelper")
    .addItem("📤 ML판매내역 → Ventas 전송", "mlSalesToVentas")
    .addItem("🔍 바코드 매칭 확인", "addBarcodeCheckToMLSheet")
    .addItem("🏭 거래처관리 팝업 열기", "inventarioToGeorae")
    .addToUi();
}

function 바코드형식통일() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    var bc = String(data[i][0]).replace(/[^0-9]/g, '');
    data[i][0] = bc ? Number(bc) : "";
  }
  sheet.getRange(2, 1, lastRow - 1, 1).setValues(data)
    .setNumberFormat("0")
    .setHorizontalAlignment("center");
  invCacheClear();
  SpreadsheetApp.getUi().alert("✅ 바코드 형식 통일 완료!");
}

// ========================
// 판매 관련
// ========================
function procesarBarcodeVentas(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var barcode = String(range.getValue()).replace(/[^0-9]/g, '');
  if (!barcode) return;

  var invData = getInvData();
  if (!invData) return;

  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][0]).replace(/[^0-9]/g, '') === barcode) {
      var cliente = sheet.getRange(row, 2).getValue().toString().toLowerCase().trim();
      var price = getMayor(cliente) ? invData[i][4] : invData[i][3];

      // 한 번에 배치로 기록 — 빠름
      sheet.getRange(row, 1, 1, 9).setValues([[
        new Date(),           // A열 날짜
        sheet.getRange(row, 2).getValue(), // B열 고객명 유지
        sheet.getRange(row, 3).getValue(), // C열 결제방법 유지
        barcode,              // D열 바코드
        invData[i][1],        // E열 SKU
        invData[i][2],        // F열 상품명
        "",                   // G열 수량 (직접입력)
        price,                // H열 단가
        ""                    // I열 금액
      ]]);

      sheet.getRange(row, 1).setNumberFormat("yyyy-MM-dd HH:mm");
      sheet.setActiveRange(sheet.getRange(row, 7)); // G열 수량으로 커서
      break;
    }
  }
}

function procesarCambioCliente(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var cliente = range.getValue().toString().toLowerCase().trim();
  var barcode = String(sheet.getRange(row, 4).getValue()).replace(/[^0-9]/g, '');
  if (!barcode) return;
  var invData = getInvData();
  if (!invData) return;
  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][0]).replace(/[^0-9]/g, '') === barcode) {
      var price = getMayor(cliente) ? invData[i][4] : invData[i][3];
      sheet.getRange(row, 8).setValue(price);
      var qty = sheet.getRange(row, 7).getValue();
      if (qty) { sheet.getRange(row, 9).setValue(qty * price); actualizarTotal(sheet); }
      break;
    }
  }
}

function procesarCantidad(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var qty = range.getValue();
  if (!qty) return;

  var barcode = String(sheet.getRange(row, 4).getValue()).replace(/[^0-9]/g, '');
  var cliente = sheet.getRange(row, 2).getValue().toString().toLowerCase().trim();
  if (!barcode) return;

  var invData = getInvData();
  if (!invData) return;

  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][0]).replace(/[^0-9]/g, '') === barcode) {
      var price = getMayor(cliente) ? invData[i][4] : invData[i][3];

      // H열 단가 + I열 금액 한번에
      sheet.getRange(row, 8, 1, 2).setValues([[price, qty * price]]);
      actualizarTotal(sheet);
      break;
    }
  }
}

function procesarPrecioManual(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var qty = sheet.getRange(row, 7).getValue();
  var price = range.getValue();
  if (qty && price) { sheet.getRange(row, 9).setValue(qty * price); actualizarTotal(sheet); }
}

function procesarBarcodeIngresos(e) {
  if (!e || !e.range) return;  // ← 이걸로 교체
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var barcode = String(range.getValue()).replace(/[^0-9]/g, '');
  if (!barcode) return;
  range.setValue(barcode);
  var invData = getInvData();
  if (!invData) return;
  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][0]).replace(/[^0-9]/g, '') === barcode) {
      if (!sheet.getRange(row, 1).getValue())
        sheet.getRange(row, 1).setValue(new Date()).setNumberFormat("yyyy-MM-dd HH:mm");
      sheet.getRange(row, 3).setValue(invData[i][1]); // SKU
      sheet.getRange(row, 4).setValue(invData[i][2]); // 상품명
      sheet.setActiveRange(sheet.getRange(row, 5));   // E열 수량으로 커서 이동
      break;
    }
  }
}

function actualizarTotal(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var values = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  var total = 0;
  for (var i = 0; i < values.length; i++)
    if (typeof values[i][0] === 'number') total += values[i][0];
  sheet.getRange(2, 10).setValue(total);
}

function registrarVentas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  var registroSheet = ss.getSheetByName(NOMBRE_REGISTRO);
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!ventasSheet || !registroSheet || !invSheet) { SpreadsheetApp.getUi().alert("시트를 찾을 수 없습니다."); return; }
  var lastRow = ventasSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("저장할 판매 데이터가 없습니다."); return; }
  var data = ventasSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 7).getValues();
  var invMap = {};
  for (var i = 0; i < invData.length; i++)
    invMap[String(invData[i][0]).replace(/[^0-9]/g, '')] = i;
  for (var i = 0; i < data.length; i++) {
    var bc = String(data[i][3]).replace(/[^0-9]/g, '');
    var qty = data[i][6];
    if (bc && qty && invMap.hasOwnProperty(bc))
      invData[invMap[bc]][6] = (Number(invData[invMap[bc]][6]) || 0) - qty;
  }
  invSheet.getRange(2, 1, invData.length, 7).setValues(invData);
  invCacheClear();
  var lastRegRow = registroSheet.getLastRow();
  registroSheet.getRange(lastRegRow + 1, 1, data.length, 10).setValues(data);
  ventasSheet.getRange(2, 1, lastRow - 1, 10).clearContent();
  SpreadsheetApp.getUi().alert(data.length + "건 판매완료!");
  통합재고동기화();
}

function cancelarVentasSeleccionadas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (!sheet) { SpreadsheetApp.getUi().alert("판매시트를 찾을 수 없습니다."); return; }
  var row = sheet.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("취소할 행을 선택하세요."); return; }
  var barcode = String(sheet.getRange(row, 4).getValue()).replace(/[^0-9]/g, '');
  var qty = sheet.getRange(row, 7).getValue();
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();
  actualizarTotal(sheet);
  if (barcode && qty) {
    actualizarStockLocal(barcode, qty, 'aumentar');
    invCacheClear();
    통합재고동기화단건(barcode, getNewStock(barcode));
  }
  SpreadsheetApp.getUi().alert("행 " + row + " 취소됨");
}

function getNewStock(barcode) {
  var invSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) return 0;
  var data = invSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++)
    if (String(data[i][0]).replace(/[^0-9]/g, '') === barcode) return data[i][6];
  return 0;
}

function actualizarStockLocal(barcode, qty, modo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) return;
  var data = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 7).getValues();
  var sb = String(barcode).replace(/[^0-9]/g, '');
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).replace(/[^0-9]/g, '') === sb) {
      invSheet.getRange(i + 2, 7).setValue(modo === 'aumentar' ? data[i][6] + qty : data[i][6] - qty);
      invCacheClear();
      return;
    }
  }
}

// ========================
// 입고 관련
// ========================
function 입고바코드확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== NOMBRE_INVENTARIO) { SpreadsheetApp.getUi().alert("Inventario 시트에서 선택 후 클릭하세요!"); return; }
  var row = sheet.getActiveCell().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("상품 행을 선택하세요!"); return; }
  var barcode = String(sheet.getRange(row, 1).getValue()).trim();
  var sku = String(sheet.getRange(row, 2).getValue()).trim();
  var producto = String(sheet.getRange(row, 3).getValue()).trim();
  var mayor = sheet.getRange(row, 5).getValue();
  if (!barcode || !sku) { SpreadsheetApp.getUi().alert("바코드 또는 SKU가 없습니다!"); return; }
  입고등록시트확인();
  var ingresos = ss.getSheetByName(NOMBRE_INGRESOS);
  if (!ingresos) { SpreadsheetApp.getUi().alert("Ingresos 시트를 찾을 수 없습니다!"); return; }
  var lastRow = ingresos.getLastRow() + 1;
  if (lastRow < 2) lastRow = 2;
  ingresos.getRange(lastRow, 1, 1, 7).setValues([[new Date(), barcode, sku, producto, "", mayor, ""]]);
  ingresos.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm");
  ss.setActiveSheet(ingresos);
  ingresos.setActiveRange(ingresos.getRange(lastRow, 5));
  입고바코드팝업(barcode, sku, producto);
}

function 입고등록시트확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("입고등록");
  if (sheet) return sheet;
  sheet = ss.insertSheet("입고등록");
  var headers = ["날짜","바코드","SKU","상품명","수량","공급가","합계","비고"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#3a2a1a").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1,35); sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,140); sheet.setColumnWidth(2,140); sheet.setColumnWidth(3,120);
  sheet.setColumnWidth(4,220); sheet.setColumnWidth(5,80); sheet.setColumnWidth(6,100);
  sheet.setColumnWidth(7,100); sheet.setColumnWidth(8,150);
  return sheet;
}

function procesarIngresos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INGRESOS);
  if (!sheet) return;

  // ← 이 부분 추가 — 현재 입력 중인 셀 값 먼저 확정
  var activeCell = sheet.getActiveCell();
  if (activeCell && activeCell.getRow() >= 2 && activeCell.getColumn() === 5) {
    var currentVal = activeCell.getValue();
    if (currentVal !== "") {
      activeCell.setValue(currentVal);
    }
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("입고 데이터가 없습니다!");
    return;
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var validData = [];
  for (var i = 0; i < data.length; i++) {
    var bc = String(data[i][1]).replace(/[^0-9]/g, '');
    var qty = Number(data[i][4]);
    if (bc && qty && qty > 0) validData.push(data[i]);
  }
  if (validData.length === 0) {
    SpreadsheetApp.getUi().alert("유효한 입고 데이터가 없습니다!\nB열 바코드와 E열 수량을 확인해주세요.");
    return;
  }
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invLastRow = invSheet.getLastRow();
  var invData = invSheet.getRange(2, 1, invLastRow - 1, 7).getValues();
  var invMap = {};
  for (var j = 0; j < invData.length; j++) {
    invMap[String(invData[j][0]).replace(/[^0-9]/g, '')] = j;
  }
  var count = 0;
  var logRows = [];
  var notFound = [];
  var updatedRows = {};

  for (var k = 0; k < validData.length; k++) {
    var barcode = String(validData[k][1]).replace(/[^0-9]/g, '');
    var cantidad = Number(validData[k][4]);
    var costo = Number(validData[k][5]);
    if (invMap.hasOwnProperty(barcode)) {
      var idx = invMap[barcode];
      var currentStock = invData[idx][6];
      currentStock = (currentStock === "" || currentStock === null || isNaN(Number(currentStock)))
                     ? 0 : Number(currentStock);
      invData[idx][6] = currentStock + cantidad;
      updatedRows[idx] = true;
      count++;
      logRows.push([
        validData[k][0] || new Date(), barcode,
        validData[k][2], validData[k][3],
        cantidad, costo, cantidad * costo, ""
      ]);
    } else {
      notFound.push(barcode);
    }
  }

  for (var idx in updatedRows) {
    var numIdx = Number(idx);
    invSheet.getRange(numIdx + 2, 7)
      .setValue(invData[numIdx][6])
      .setNumberFormat("0")
      .setFontColor("#000000")
      .setHorizontalAlignment("center");
  }

  invCacheClear();

  if (logRows.length > 0) {
    var 등록Sheet = 입고등록시트확인();
    var 등록LastRow = 등록Sheet.getLastRow() + 1;
    등록Sheet.getRange(등록LastRow, 1, logRows.length, 8).setValues(logRows);
    등록Sheet.getRange(등록LastRow, 1, logRows.length, 1).setNumberFormat("yyyy-MM-dd HH:mm");
    등록Sheet.getRange(등록LastRow, 6, logRows.length, 2).setNumberFormat("#,##0");
  }

  sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  통합재고동기화();

  var msg = "✅ 입고완료!\n" + count + "개 상품 재고 업데이트!";
  if (notFound.length > 0) msg += "\n\n❌ 바코드 없음: " + notFound.join(", ");
  SpreadsheetApp.getUi().alert(msg);
}

function 입고바코드팝업(barcode, sku, producto) {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:16px;text-align:center;background:#f0f7f0;">' +
    '<div style="background:#1a3a2a;color:white;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;font-weight:bold;">📦 입고 바코드</div>' +
    '<div style="background:white;border-radius:8px;padding:10px;margin-bottom:10px;">' +
    '<img src="https://barcode.tec-it.com/barcode.ashx?data=' + barcode + '&code=Code128&dpi=96&imagetype=png" style="width:100%;max-width:220px;" /></div>' +
    '<div style="background:white;border-radius:8px;padding:10px;font-size:12px;text-align:left;">' +
    '<div style="margin-bottom:6px;"><span style="color:#888;">바코드</span><br><strong>' + barcode + '</strong></div>' +
    '<div style="margin-bottom:6px;"><span style="color:#888;">SKU</span><br><strong>' + sku + '</strong></div>' +
    '<div><span style="color:#888;">상품명</span><br><strong>' + producto + '</strong></div></div>' +
    '<div style="margin-top:12px;background:#e8f5e9;border-radius:8px;padding:10px;font-size:12px;color:#1a3a2a;font-weight:bold;">' +
    '✅ E열에 수량을 입력하세요!</div></div>'
  ).setTitle("📦 입고 바코드").setWidth(260);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ========================
// 통합재고 동기화
// ========================
function 통합재고동기화() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var inv = ss.getSheetByName(NOMBRE_INVENTARIO);
    if (!inv) return;
    var invData = inv.getDataRange().getValues();
    var stockMap = {};
    for (var i = 1; i < invData.length; i++) {
      var bc = String(invData[i][0]).trim();
      if (bc) stockMap[bc] = Number(invData[i][6]) || 0;
    }
    var 통합ss = SpreadsheetApp.openById(ID_통합재고);
    var 통합Sheet = 통합ss.getSheets()[0];
    var isNAOS = ss.getName().toUpperCase().indexOf("NAOS") !== -1;
    var stockCol = isNAOS ? 4 : 5;
    var DS = 4;
    var lastRow = 통합Sheet.getLastRow();
    var 통합Data = 통합Sheet.getRange(DS, 1, lastRow-DS+1, stockCol).getValues();
    for (var j = 0; j < 통합Data.length; j++) {
      var bc = String(통합Data[j][0]).trim();
      if (bc && stockMap[bc] !== undefined) 통합Data[j][stockCol-1] = stockMap[bc];
    }
    통합Sheet.getRange(DS, 1, lastRow-DS+1, stockCol).setValues(통합Data);
  } catch(e) { console.log("동기화 오류: " + e.message); }
}

function 통합재고동기화단건(barcode, newStock) {
  try {
    var 통합ss = SpreadsheetApp.openById(ID_통합재고);
    var 통합Sheet = 통합ss.getSheets()[0];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stockCol = ss.getName().toUpperCase().indexOf("NAOS") !== -1 ? 4 : 5;
    var DS = 4;
    var barcodes = 통합Sheet.getRange(DS, 1, 통합Sheet.getLastRow()-DS+1, 1).getValues();
    for (var j = 0; j < barcodes.length; j++) {
      if (String(barcodes[j][0]).trim() === String(barcode).trim()) {
        통합Sheet.getRange(DS+j, stockCol).setValue(newStock); break;
      }
    }
  } catch(e) { console.log("단건동기화 오류: " + e.message); }
}

function 통합재고자동동기화() {
  var hour = new Date().getHours();
  if (hour < 9 || hour > 20) return;
  통합재고동기화();
}

// ========================
// 부족재고
// ========================
function 부족재고표시() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) return;
  var data = invSheet.getDataRange().getValues();
  var lowStock = [];
  for (var i = 1; i < data.length; i++) {
    var sku = String(data[i][1]).trim();
    var stock = Number(data[i][6]);
    if (sku && stock < 10) lowStock.push([sku, stock]);
  }
  lowStock.sort(function(a,b){return a[1]-b[1];});
  var startRow = 2;
  var lastRow = invSheet.getLastRow();
  if (lastRow >= startRow) invSheet.getRange(startRow, 9, lastRow, 2).clearContent();
  if (lowStock.length === 0) { invSheet.getRange(startRow, 9).setValue("부족재고 없음"); return; }
  invSheet.getRange(1,9).setValue("⚠️ 부족재고 SKU").setBackground("#E74C3C").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  invSheet.getRange(1,10).setValue("재고").setBackground("#E74C3C").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  for (var j = 0; j < lowStock.length; j++) {
    var r = startRow + j;
    var bg = lowStock[j][1] === 0 ? "#FADBD8" : "#FEF9E7";
    invSheet.getRange(r,9).setValue(lowStock[j][0]).setBackground(bg).setHorizontalAlignment("center");
    invSheet.getRange(r,10).setValue(lowStock[j][1]).setBackground(bg).setHorizontalAlignment("center").setFontWeight("bold").setFontColor(lowStock[j][1]===0?"#E74C3C":"#E67E22");
  }
}

// ========================
// 바코드 팝업
// ========================
function 바코드팝업(barcode, sku, producto) {
  var html = HtmlService.createHtmlOutput(
    '<html><body style="text-align:center;font-family:Arial;padding:15px;margin:0;background:#f8f9fa;">' +
    '<div style="background:white;border-radius:10px;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,0.1);">' +
    '<h3 style="color:#2C3E50;margin:0 0 5px 0;font-size:16px;">' + sku + '</h3>' +
    '<p style="color:#555;font-size:12px;margin:5px 0;">' + producto + '</p>' +
    '<p style="color:#888;font-size:11px;margin:5px 0;border-bottom:1px solid #eee;padding-bottom:10px;">바코드: ' + barcode + '</p>' +
    '<img src="https://barcode.tec-it.com/barcode.ashx?data=' + barcode + '&code=Code128&dpi=96&imagetype=png" ' +
    'style="width:100%;max-width:280px;height:100px;margin:15px 0;border:1px solid #ddd;padding:8px;border-radius:5px;" />' +
    '</div></body></html>'
  ).setTitle('🔍 바코드 확인').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var row = e.range.getRow();
    if (sheet.getName() === NOMBRE_INVENTARIO && row >= 2) {
      var barcode = sheet.getRange(row, 1).getValue();
      var sku = sheet.getRange(row, 2).getValue();
      var producto = sheet.getRange(row, 3).getValue();
      if (barcode && sku) 바코드팝업(String(barcode), String(sku), String(producto));
    }
  } catch(err) { console.log("onSelectionChange 오류: " + err.toString()); }
}

function 선택셀바코드확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("상품 행을 선택하세요!"); return; }
  var barcode = String(sheet.getRange(row, 1).getValue()).trim();
  var sku = String(sheet.getRange(row, 2).getValue()).trim();
  var producto = String(sheet.getRange(row, 3).getValue()).trim();
  var precioDetalle = sheet.getRange(row, 4).getValue();
  if (!barcode) { SpreadsheetApp.getUi().alert("바코드가 없습니다!"); return; }
  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (!ventasSheet) { SpreadsheetApp.getUi().alert("Ventas 시트를 찾을 수 없습니다!"); return; }
  ss.setActiveSheet(ventasSheet);
  var lastRow = ventasSheet.getLastRow() + 1;
  if (lastRow < 2) lastRow = 2;
  ventasSheet.getRange(lastRow, 1, 1, 10).setValues([[new Date(),"","",barcode,sku,producto,"",precioDetalle,"",""]]);
  ventasSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm");
  ventasSheet.setActiveRange(ventasSheet.getRange(lastRow, 7));
  바코드팝업(barcode, sku, producto);
}

function 바코드직접생성() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) { SpreadsheetApp.getUi().alert("Inventario 시트를 찾을 수 없습니다!"); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow-1, 2).getValues();
  var maxBarcode = 2000000000000;
  for (var i = 0; i < data.length; i++) {
    var bc = Number(String(data[i][0]).replace(/[^0-9]/g,''));
    if (bc > maxBarcode) maxBarcode = bc;
  }
  var count = 0;
  for (var i = 0; i < data.length; i++)
    if (!data[i][0] && data[i][1]) { maxBarcode++; data[i][0] = maxBarcode; count++; }
  if (count === 0) { SpreadsheetApp.getUi().alert("바코드가 없는 SKU가 없습니다!"); return; }
  sheet.getRange(2, 1, data.length, 1).setValues(data.map(function(r){return [r[0]];}));
  invCacheClear();
  SpreadsheetApp.getUi().alert("✅ " + count + "개 바코드 생성완료!");
}

function 시트디자인적용() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [
    {name:"Inventario",headerColor:"#1a3a2a",oddColor:"#f0f7f0",evenColor:"#ffffff",lastCol:7},
    {name:"Ventas",headerColor:"#1a2a3a",oddColor:"#f0f4f7",evenColor:"#ffffff",lastCol:10},
    {name:"Ingresos",headerColor:"#3a2a1a",oddColor:"#f7f4f0",evenColor:"#ffffff",lastCol:7},
    {name:"Registro de ventas",headerColor:"#2a1a3a",oddColor:"#f4f0f7",evenColor:"#ffffff",lastCol:10},
    {name:"입고등록",headerColor:"#3a2a1a",oddColor:"#f7f4f0",evenColor:"#ffffff",lastCol:8},
    {name:"재고수동조정",headerColor:"#1a2a3a",oddColor:"#f0f4f7",evenColor:"#ffffff",lastCol:9}
  ];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = ss.getSheetByName(sheets[s].name);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    var lastCol = sheets[s].lastCol;
    if (lastRow < 1) continue;
    sheet.getRange(1,1,1,lastCol).setBackground(sheets[s].headerColor).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(1,35);
    if (lastRow >= 2)
      for (var i = 2; i <= lastRow; i++)
        sheet.getRange(i,1,1,lastCol).setBackground(i%2===0?sheets[s].evenColor:sheets[s].oddColor);
    sheet.getRange(1,1,lastRow,lastCol).setBorder(true,true,true,true,true,true,"#cccccc",SpreadsheetApp.BorderStyle.SOLID);
  }
  SpreadsheetApp.getUi().alert("✅ 디자인 적용 완료!");
}

function 행색상자동적용(sheet, row) {
  var colorMap = {
    "Inventario":{odd:"#f0f7f0",even:"#ffffff",lastCol:7},
    "Ventas":{odd:"#f0f4f7",even:"#ffffff",lastCol:10},
    "Ingresos":{odd:"#f7f4f0",even:"#ffffff",lastCol:7},
    "Registro de ventas":{odd:"#f4f0f7",even:"#ffffff",lastCol:10}
  };
  var colors = colorMap[sheet.getName()];
  if (!colors) return;
  sheet.getRange(row,1,1,colors.lastCol).setBackground(row%2===0?colors.even:colors.odd);
}

// ========================
// 재고수동조정
// ========================
function 재고수동조정시트확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("재고수동조정");
  if (sheet) return sheet;
  sheet = ss.insertSheet("재고수동조정");
  var headers = ["날짜","바코드","SKU","상품명","기존재고","조정수량(+/-)","새재고","이유","처리결과"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1,35); sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,140); sheet.setColumnWidth(2,140); sheet.setColumnWidth(3,120);
  sheet.setColumnWidth(4,220); sheet.setColumnWidth(5,80); sheet.setColumnWidth(6,120);
  sheet.setColumnWidth(7,80); sheet.setColumnWidth(8,200); sheet.setColumnWidth(9,120);
  return sheet;
}

function 재고수동조정처리() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("재고수동조정");
  if (!sheet) { SpreadsheetApp.getUi().alert("재고수동조정 시트가 없습니다!"); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("조정 데이터가 없습니다!"); return; }
  var data = sheet.getRange(2,1,lastRow-1,9).getValues();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) return;
  var invData = invSheet.getRange(2,1,invSheet.getLastRow()-1,7).getValues();
  var invMap = {};
  for (var i = 0; i < invData.length; i++) {
    invMap[String(invData[i][0]).replace(/[^0-9]/g, '')] = i; // ← 수정
  }
  var count = 0;
  for (var j = 0; j < data.length; j++) {
    if (data[j][8] && data[j][8] !== "") continue;
    var barcode = String(data[j][1]).replace(/[^0-9]/g, ''); // ← 수정
    var qty = Number(data[j][5]);
    if (!barcode || isNaN(qty) || qty === 0) continue;
    if (invMap.hasOwnProperty(barcode)) {
      var idx = invMap[barcode];
      var cur = Number(invData[idx][6]) || 0;
      var newS = Math.max(0, cur + qty);
      invData[idx][6] = newS;
      sheet.getRange(j+2,5).setValue(cur).setBackground("#f5f5f5").setFontColor("#888888");
      sheet.getRange(j+2,7).setValue(newS).setBackground("#e8f5e9").setFontColor("#1a3a2a").setFontWeight("bold");
      sheet.getRange(j+2,9).setValue("✅ "+cur+" → "+newS).setBackground("#e8f5e9").setFontColor("#1a3a2a").setFontWeight("bold");
      if (!data[j][0]) sheet.getRange(j+2,1).setValue(new Date()).setNumberFormat("yyyy-MM-dd HH:mm");
      count++;
    } else {
      sheet.getRange(j+2,9).setValue("❌ 바코드 없음").setBackground("#fadbd8").setFontColor("#e74c3c").setFontWeight("bold");
    }
  }
  invSheet.getRange(2,1,invData.length,7).setValues(invData);
  invCacheClear();
  통합재고동기화();
  SpreadsheetApp.getUi().alert("✅ 재고조정 완료!\n" + count + "개 상품 업데이트!");
}

function 재고수동조정열기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = 재고수동조정시트확인();
  ss.setActiveSheet(sheet);
  var lastRow = sheet.getLastRow() + 1;
  if (lastRow < 2) lastRow = 2;
  sheet.setActiveRange(sheet.getRange(lastRow, 2));
  SpreadsheetApp.getUi().alert("📋 재고수동조정\n\n① B열: 바코드 스캔\n② E열: 기존재고 자동표시\n③ F열: 조정수량 (+10 또는 -5)\n④ G열: 새재고 자동표시\n⑤ H열: 이유 입력\n⑥ 재고조정 반영 클릭");
}

// ========================
// 매출분석
// ========================
function 매출분석업데이트() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var registroSheet = ss.getSheetByName(NOMBRE_REGISTRO);
  if (!registroSheet || registroSheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert("판매기록 데이터가 없습니다!"); return; }
  var lastRow = registroSheet.getLastRow();
  var data = registroSheet.getRange(2,1,lastRow-1,10).getValues();
  data = data.filter(function(r){return r[0]&&r[8];});
  if (data.length === 0) { SpreadsheetApp.getUi().alert("유효한 판매 데이터가 없습니다!"); return; }
  var analysisSheet = ss.getSheetByName("매출분석");
  if (!analysisSheet) analysisSheet = ss.insertSheet("매출분석");
  else { analysisSheet.clearContents(); analysisSheet.clearFormats(); }
  var writeRow = 1;
  writeRow = 매출분석_일별(analysisSheet, data, writeRow); writeRow += 2;
  writeRow = 매출분석_월별(analysisSheet, data, writeRow); writeRow += 2;
  writeRow = 매출분석_년별(analysisSheet, data, writeRow); writeRow += 2;
  writeRow = 매출분석_거래처별(analysisSheet, data, writeRow); writeRow += 2;
  writeRow = 매출분석_상품별(analysisSheet, data, writeRow);
  analysisSheet.setColumnWidth(1,150); analysisSheet.setColumnWidth(2,120);
  analysisSheet.setColumnWidth(3,120); analysisSheet.setColumnWidth(4,120);
  ss.setActiveSheet(analysisSheet);
  SpreadsheetApp.getUi().alert("✅ 매출분석 업데이트 완료!");
}

function 매출분석_일별(sheet, data, startRow) {
  sheet.getRange(startRow,1,1,4).setValues([["📅 일별 판매합계","","",""]]).setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  sheet.getRange(startRow,1,1,4).merge(); startRow++;
  sheet.getRange(startRow,1,1,4).setValues([["날짜","판매건수","판매금액","평균단가"]]).setBackground("#2c3e50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center"); startRow++;
  var map={};
  for(var i=0;i<data.length;i++){var key=Utilities.formatDate(new Date(data[i][0]),"America/Santiago","yyyy-MM-dd");var amt=Number(data[i][8])||0;if(!map[key])map[key]={count:0,total:0};map[key].count++;map[key].total+=amt;}
  var keys=Object.keys(map).sort().reverse();
  var rows=keys.map(function(k){return[k,map[k].count,map[k].total,map[k].count>0?Math.round(map[k].total/map[k].count):0];});
  if(rows.length>0){sheet.getRange(startRow,1,rows.length,4).setValues(rows);sheet.getRange(startRow,3,rows.length,2).setNumberFormat("#,##0");for(var r=0;r<rows.length;r++)sheet.getRange(startRow+r,1,1,4).setBackground(r%2===0?"#f0f4f7":"#ffffff");var tr=startRow+rows.length;sheet.getRange(tr,1,1,4).setValues([["합계",rows.length+"일",rows.reduce(function(s,r){return s+r[2];},0),""]]).setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold");sheet.getRange(tr,3).setNumberFormat("#,##0");startRow=tr+1;}
  return startRow;
}

function 매출분석_월별(sheet, data, startRow) {
  sheet.getRange(startRow,1,1,4).setValues([["📆 월별 판매합계","","",""]]).setBackground("#1a3a2a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  sheet.getRange(startRow,1,1,4).merge(); startRow++;
  sheet.getRange(startRow,1,1,4).setValues([["년월","판매건수","판매금액","전월대비"]]).setBackground("#2c3e50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center"); startRow++;
  var map={};
  for(var i=0;i<data.length;i++){var key=Utilities.formatDate(new Date(data[i][0]),"America/Santiago","yyyy-MM");var amt=Number(data[i][8])||0;if(!map[key])map[key]={count:0,total:0};map[key].count++;map[key].total+=amt;}
  var keys=Object.keys(map).sort().reverse();
  var rows=keys.map(function(k,idx){var prev=keys[idx+1]?map[keys[idx+1]].total:0;var diff=prev>0?Math.round((map[k].total-prev)/prev*100):0;var ds=prev>0?(diff>0?"▲"+diff+"%":"▼"+Math.abs(diff)+"%"):"-";return[k,map[k].count,map[k].total,ds];});
  if(rows.length>0){sheet.getRange(startRow,1,rows.length,4).setValues(rows);sheet.getRange(startRow,3,rows.length,1).setNumberFormat("#,##0");for(var r=0;r<rows.length;r++){sheet.getRange(startRow+r,1,1,4).setBackground(r%2===0?"#f0f7f0":"#ffffff");var dv=rows[r][3];if(String(dv).indexOf("▲")!==-1)sheet.getRange(startRow+r,4).setFontColor("#27ae60").setFontWeight("bold");else if(String(dv).indexOf("▼")!==-1)sheet.getRange(startRow+r,4).setFontColor("#e74c3c").setFontWeight("bold");}var tr=startRow+rows.length;sheet.getRange(tr,1,1,4).setValues([["합계",rows.reduce(function(s,r){return s+r[1];},0)+"건",rows.reduce(function(s,r){return s+r[2];},0),""]]).setBackground("#1a3a2a").setFontColor("#FFFFFF").setFontWeight("bold");sheet.getRange(tr,3).setNumberFormat("#,##0");startRow=tr+1;}
  return startRow;
}

function 매출분석_년별(sheet, data, startRow) {
  sheet.getRange(startRow,1,1,4).setValues([["📊 년별 판매합계","","",""]]).setBackground("#3a1a2a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  sheet.getRange(startRow,1,1,4).merge(); startRow++;
  sheet.getRange(startRow,1,1,4).setValues([["년도","판매건수","판매금액","전년대비"]]).setBackground("#2c3e50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center"); startRow++;
  var map={};
  for(var i=0;i<data.length;i++){var key=Utilities.formatDate(new Date(data[i][0]),"America/Santiago","yyyy");var amt=Number(data[i][8])||0;if(!map[key])map[key]={count:0,total:0};map[key].count++;map[key].total+=amt;}
  var keys=Object.keys(map).sort().reverse();
  var rows=keys.map(function(k,idx){var prev=keys[idx+1]?map[keys[idx+1]].total:0;var diff=prev>0?Math.round((map[k].total-prev)/prev*100):0;var ds=prev>0?(diff>0?"▲"+diff+"%":"▼"+Math.abs(diff)+"%"):"-";return[k,map[k].count,map[k].total,ds];});
  if(rows.length>0){sheet.getRange(startRow,1,rows.length,4).setValues(rows);sheet.getRange(startRow,3,rows.length,1).setNumberFormat("#,##0");for(var r=0;r<rows.length;r++){sheet.getRange(startRow+r,1,1,4).setBackground(r%2===0?"#f4f0f7":"#ffffff");var dv=rows[r][3];if(String(dv).indexOf("▲")!==-1)sheet.getRange(startRow+r,4).setFontColor("#27ae60").setFontWeight("bold");else if(String(dv).indexOf("▼")!==-1)sheet.getRange(startRow+r,4).setFontColor("#e74c3c").setFontWeight("bold");}var tr=startRow+rows.length;sheet.getRange(tr,1,1,4).setValues([["합계",rows.reduce(function(s,r){return s+r[1];},0)+"건",rows.reduce(function(s,r){return s+r[2];},0),""]]).setBackground("#3a1a2a").setFontColor("#FFFFFF").setFontWeight("bold");sheet.getRange(tr,3).setNumberFormat("#,##0");startRow=tr+1;}
  return startRow;
}

function 매출분석_거래처별(sheet, data, startRow) {
  sheet.getRange(startRow,1,1,4).setValues([["🏢 거래처별 매출","","",""]]).setBackground("#3a2a1a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  sheet.getRange(startRow,1,1,4).merge(); startRow++;
  sheet.getRange(startRow,1,1,4).setValues([["거래처","판매건수","판매금액","비율"]]).setBackground("#2c3e50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center"); startRow++;
  var map={};var grandTotal=0;
  for(var i=0;i<data.length;i++){var c=String(data[i][1]||"미입력").trim();var amt=Number(data[i][8])||0;if(!map[c])map[c]={count:0,total:0};map[c].count++;map[c].total+=amt;grandTotal+=amt;}
  var clients=Object.keys(map).sort(function(a,b){return map[b].total-map[a].total;}).slice(0,10);
  var rows=clients.map(function(c){return[c,map[c].count,map[c].total,grandTotal>0?(map[c].total/grandTotal*100).toFixed(1)+"%":"0%"];});
  if(rows.length>0){sheet.getRange(startRow,1,rows.length,4).setValues(rows);sheet.getRange(startRow,3,rows.length,1).setNumberFormat("#,##0");for(var r=0;r<rows.length;r++){sheet.getRange(startRow+r,1,1,4).setBackground(r%2===0?"#f7f4f0":"#ffffff");if(r===0)sheet.getRange(startRow+r,1,1,4).setBackground("#ffd700").setFontWeight("bold");if(r===1)sheet.getRange(startRow+r,1,1,4).setBackground("#c0c0c0").setFontWeight("bold");if(r===2)sheet.getRange(startRow+r,1,1,4).setBackground("#cd7f32").setFontWeight("bold");}var tr=startRow+rows.length;sheet.getRange(tr,1,1,4).setValues([["합계",rows.reduce(function(s,r){return s+r[1];},0)+"건",grandTotal,"100%"]]).setBackground("#3a2a1a").setFontColor("#FFFFFF").setFontWeight("bold");sheet.getRange(tr,3).setNumberFormat("#,##0");startRow=tr+1;}
  return startRow;
}

function 매출분석_상품별(sheet, data, startRow) {
  sheet.getRange(startRow,1,1,5).setValues([["🏆 상품별 판매순위","","","",""]]).setBackground("#2a1a3a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  sheet.getRange(startRow,1,1,5).merge(); startRow++;
  sheet.getRange(startRow,1,1,5).setValues([["순위","SKU","상품명","판매수량","판매금액"]]).setBackground("#2c3e50").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setColumnWidth(5,130); startRow++;
  var map={};
  for(var i=0;i<data.length;i++){var sku=String(data[i][4]||"").trim();if(!sku)continue;if(!map[sku])map[sku]={nombre:String(data[i][5]||"").trim(),qty:0,total:0};map[sku].qty+=Number(data[i][6])||0;map[sku].total+=Number(data[i][8])||0;}
  var skus=Object.keys(map).sort(function(a,b){return map[b].total-map[a].total;}).slice(0,10);
  var rows=skus.map(function(s,p){return[p+1,s,map[s].nombre,map[s].qty,map[s].total];});
  if(rows.length>0){sheet.getRange(startRow,1,rows.length,5).setValues(rows);sheet.getRange(startRow,5,rows.length,1).setNumberFormat("#,##0");for(var r=0;r<rows.length;r++){sheet.getRange(startRow+r,1,1,5).setBackground(r%2===0?"#f4f0f7":"#ffffff");if(r===0){sheet.getRange(startRow+r,1,1,5).setBackground("#ffd700").setFontWeight("bold");sheet.getRange(startRow+r,1).setValue("🥇 1");}if(r===1){sheet.getRange(startRow+r,1,1,5).setBackground("#c0c0c0").setFontWeight("bold");sheet.getRange(startRow+r,1).setValue("🥈 2");}if(r===2){sheet.getRange(startRow+r,1,1,5).setBackground("#cd7f32").setFontWeight("bold");sheet.getRange(startRow+r,1).setValue("🥉 3");}}var tr=startRow+rows.length;sheet.getRange(tr,1,1,5).setValues([["합계","","",rows.reduce(function(s,r){return s+r[3];},0),rows.reduce(function(s,r){return s+r[4];},0)]]).setBackground("#2a1a3a").setFontColor("#FFFFFF").setFontWeight("bold");sheet.getRange(tr,4,1,2).setNumberFormat("#,##0");startRow=tr+1;}
  return startRow;
}

// ========================
// 상품검색 사이드바
// ========================
function 상품검색사이드바() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (invSheet) ss.setActiveSheet(invSheet);
  var data = getInventarioData();
  var htmlContent = 상품검색HTML생성(JSON.stringify(data));
  var html = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('🔍 Buscar / 상품 검색')
    .setWidth(280);
  SpreadsheetApp.getUi().showSidebar(html);
}

function 상품검색HTML생성(dataJson) {
  return '<!DOCTYPE html><html><head><style>' +
    '*{box-sizing:border-box;margin:0;padding:0;}' +
    'body{font-family:Arial,sans-serif;background:#f0f4f7;padding:12px;font-size:13px;}' +
    '.header{background:#1a2a3a;color:white;padding:10px 12px;border-radius:8px;margin-bottom:10px;font-weight:bold;font-size:14px;text-align:center;}' +
    '.search-box{display:flex;gap:6px;margin-bottom:6px;}' +
    'input[type=text]{flex:1;padding:9px 10px;border:2px solid #1a2a3a;border-radius:6px;font-size:13px;outline:none;}' +
    '.btn{padding:8px 12px;background:#1a2a3a;color:white;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:bold;}' +
    '.btn-clear{width:100%;padding:8px;background:#e74c3c;color:white;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:bold;margin-bottom:8px;display:none;}' +
    '.scan-btn{width:100%;padding:9px;background:#1a3a2a;color:white;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:bold;margin-bottom:8px;}' +
    '.scan-box{display:none;margin-bottom:8px;}' +
    '.scan-box input{width:100%;padding:9px 10px;border:2px solid #1a3a2a;border-radius:6px;background:#e8f5e9;font-size:13px;outline:none;}' +
    '.status{font-size:11px;color:#888;margin-bottom:6px;text-align:right;}' +
    '.results{background:white;border-radius:8px;overflow-y:auto;max-height:340px;box-shadow:0 2px 6px rgba(0,0,0,0.1);}' +
    '.result-item{padding:9px 12px;border-bottom:1px solid #eee;cursor:pointer;display:flex;align-items:center;gap:8px;}' +
    '.result-item:hover{background:#e8f4fd;}' +
    '.item-info{flex:1;min-width:0;}' +
    '.sku{font-weight:bold;color:#1a2a3a;font-size:12px;}' +
    '.nombre{color:#555;font-size:11px;margin:2px 0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}' +
    '.info-row{display:flex;gap:4px;margin-top:3px;flex-wrap:wrap;}' +
    '.badge{font-size:10px;padding:2px 5px;border-radius:4px;font-weight:bold;}' +
    '.stock-ok{background:#e8f5e9;color:#1a3a2a;}' +
    '.stock-low{background:#fef9e7;color:#e67e22;}' +
    '.stock-zero{background:#fadbd8;color:#e74c3c;}' +
    '.price{background:#eaf0fb;color:#2980b9;}' +
    '.no-result{padding:24px;text-align:center;color:#aaa;font-size:12px;line-height:1.6;}' +
    '.tip{font-size:11px;color:#aaa;text-align:center;margin-top:8px;line-height:1.6;}' +
    'hr{border:none;border-top:1px solid #ddd;margin:8px 0;}' +
    '.overlay{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999;display:flex;align-items:center;justify-content:center;}' +
    '.popup{background:white;border-radius:12px;padding:16px;width:248px;max-height:92vh;overflow-y:auto;box-shadow:0 10px 30px rgba(0,0,0,0.3);}' +
    '.popup-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;}' +
    '.popup-title{font-weight:bold;font-size:13px;color:#1a2a3a;}' +
    '.popup-close{background:#e74c3c;color:white;border:none;border-radius:50%;width:24px;height:24px;cursor:pointer;font-size:12px;}' +
    '.popup-info{background:#f8f9fa;border-radius:8px;padding:10px;margin-bottom:8px;}' +
    '.popup-sku{font-size:14px;font-weight:bold;color:#1a2a3a;margin-bottom:4px;}' +
    '.popup-nombre{font-size:11px;color:#555;line-height:1.4;}' +
    '.popup-prices{display:flex;gap:6px;margin-bottom:8px;}' +
    '.popup-price{flex:1;border-radius:8px;padding:8px;text-align:center;}' +
    '.popup-price-label{font-size:10px;margin-bottom:2px;}' +
    '.popup-price-value{font-size:13px;font-weight:bold;color:#1a2a3a;}' +
    '.popup-stock{border-radius:8px;padding:8px;text-align:center;margin-bottom:8px;font-size:12px;font-weight:bold;}' +
    '.popup-barcode-num{font-size:12px;color:#555;font-weight:bold;padding:8px;text-align:center;margin-bottom:10px;background:#f8f9fa;border-radius:8px;}' +
    '.popup-btn{width:100%;padding:10px;background:#1a2a3a;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:bold;margin-bottom:6px;}' +
    '</style></head><body>' +
    '<div id="popupContainer"></div>' +
    '<div class="header">🔍 Buscar / 상품 검색</div>' +
    '<button class="scan-btn" onclick="toggleScan()">📷 Escanear / 바코드 스캔</button>' +
    '<div class="scan-box" id="scanBox">' +
    '<input type="text" id="scanInput" placeholder="Escanear codigo..." oninput="onScanInput(this.value)" /></div>' +
    '<hr>' +
    '<div class="search-box">' +
    '<input type="text" id="searchInput" placeholder="SKU o nombre..." oninput="onSearchInput(this.value)" autofocus />' +
    '<button class="btn" onclick="doSearch()">🔍</button></div>' +
    '<button class="btn-clear" id="clearBtn" onclick="clearAll()">✖ Mostrar todo / 전체보기</button>' +
    '<div class="status" id="statusLabel">✅ 준비완료</div>' +
    '<div class="results" id="results">' +
    '<div class="no-result">Ingrese texto para buscar<br>검색어를 입력하세요</div></div>' +
    '<div class="tip">💡 Clic para ver detalle / 클릭시 상세정보</div>' +
    '<script>' +
    'var allData=' + dataJson + ';' +
    'var searchTimer=null;' +
    'var currentFiltered=[];' +
    'var showLimit=100;' +
    'document.getElementById("statusLabel").textContent="✅ "+allData.length+"개 준비완료";' +
    'function toggleScan(){' +
    'var box=document.getElementById("scanBox");' +
    'var input=document.getElementById("scanInput");' +
    'if(!box.style.display||box.style.display==="none"){' +
    'box.style.display="block";setTimeout(function(){input.focus();},100);}' +
    'else{box.style.display="none";input.value="";}}' +
    'function onScanInput(value){' +
    'var barcode=value.replace(/[^0-9]/g,"");' +
    'if(barcode.length>=8){' +
    'var found=allData.filter(function(item){return String(item[0]).trim()===barcode;});' +
    'if(found.length>0){currentFiltered=found;renderResults(found,100);showPopup(0);}' +
    'else{document.getElementById("results").innerHTML="<div class=no-result>❌ No encontrado<br>"+barcode+"</div>";}' +
    'document.getElementById("scanInput").value="";}}' +
    'function onSearchInput(value){' +
    'clearTimeout(searchTimer);' +
    'if(!value||value.trim()===""){clearAll();return;}' +
    'searchTimer=setTimeout(function(){doSearch();},100);}' +
    'function doSearch(){' +
    'var query=document.getElementById("searchInput").value.trim();' +
    'if(!query){clearAll();return;}' +
    'var q=query.toLowerCase();' +
    'currentFiltered=allData.filter(function(item){' +
    'return String(item[0]).toLowerCase().indexOf(q)!==-1||' +
    'String(item[1]).toLowerCase().indexOf(q)!==-1||' +
    'String(item[2]).toLowerCase().indexOf(q)!==-1;});' +
    'showLimit=100;' +
    'renderResults(currentFiltered,showLimit);' +
    'var msg=currentFiltered.length+"개 검색됨";' +
    'if(currentFiltered.length>100) msg+=" — 더보기 ▼";' +
    'document.getElementById("statusLabel").textContent=msg;' +
    'document.getElementById("clearBtn").style.display="block";}' +
    'function renderResults(filtered,limit){' +
    'var container=document.getElementById("results");' +
    'if(!filtered||filtered.length===0){' +
    'container.innerHTML="<div class=no-result>No se encontraron resultados</div>";return;}' +
    'var html="";' +
    'var count=Math.min(filtered.length,limit);' +
    'for(var i=0;i<count;i++){' +
    'var item=filtered[i];' +
    'var sku=item[1]||"";' +
    'var nombre=item[2]||"";' +
    'var detalle=Number(item[3]||0).toLocaleString();' +
    'var mayor=Number(item[4]||0).toLocaleString();' +
    'var stock=Number(item[6]||0);' +
    'var sc=stock===0?"stock-zero":stock<10?"stock-low":"stock-ok";' +
    'var st=stock===0?"Agotado/품절":"Stock: "+stock;' +
    'html+="<div class=result-item onclick=\'showPopup("+i+")\'>"+' +
    '"<div class=item-info>"+' +
    '"<div class=sku>"+sku+"</div>"+' +
    '"<div class=nombre>"+nombre+"</div>"+' +
    '"<div class=info-row>"+' +
    '"<span class=\'badge "+sc+"\'>"+st+"</span>"+' +
    '"<span class=\'badge price\'>$"+detalle+"</span>"+' +
    '"<span class=\'badge price\'>M $"+mayor+"</span>"+' +
    '"</div></div>"+' +
    '"<span style=\'color:#aaa;font-size:16px\'>›</span></div>";}' +
    'if(filtered.length>limit){' +
    'html+="<div style=\'padding:12px;text-align:center;cursor:pointer;color:#2980b9;font-weight:bold;background:#eaf0fb;border-radius:8px;margin:4px;\' onclick=\'showMore()\'>"+(filtered.length-limit)+" más... / 더보기 👇</div>";}' +
    'container.innerHTML=html;}' +
    'function showMore(){' +
    'showLimit+=100;' +
    'renderResults(currentFiltered,showLimit);' +
    'document.getElementById("statusLabel").textContent=' +
    'currentFiltered.length+"개 중 "+Math.min(showLimit,currentFiltered.length)+"개 표시중";}' +
    'function showPopup(idx){' +
    'var item=currentFiltered[idx];if(!item)return;' +
    'var bc=item[0]||"";' +
    'var sku=item[1]||"";' +
    'var nombre=item[2]||"";' +
    'var detalle=Number(item[3]||0).toLocaleString();' +
    'var mayor=Number(item[4]||0).toLocaleString();' +
    'var rowNum=Number(item[5]||0);' +
    'var stock=Number(item[6]||0);' +
    'var stockColor=stock===0?"#e74c3c":stock<10?"#e67e22":"#27ae60";' +
    'var stockBg=stock===0?"#fadbd8":stock<10?"#fef9e7":"#e8f5e9";' +
    'var stockText=stock===0?"❌ Agotado / 품절":"✅ Stock: "+stock+"개";' +
    'var popup=' +
    '"<div class=\'overlay\' onclick=\'if(event.target.className===\\\"overlay\\\")closePopup()\'>"+' +
    '"<div class=\'popup\'>"+' +
    '"<div class=\'popup-header\'>"+' +
    '"<span class=\'popup-title\'>📦 Detalle Producto</span>"+' +
    '"<button class=\'popup-close\' onclick=\'closePopup()\'>✕</button></div>"+' +
    '"<div class=\'popup-info\'>"+' +
    '"<div class=\'popup-sku\'>"+sku+"</div>"+' +
    '"<div class=\'popup-nombre\'>"+nombre+"</div></div>"+' +
    '"<div class=\'popup-prices\'>"+' +
    '"<div class=\'popup-price\' style=\'background:#eaf0fb\'>"+' +
    '"<div class=\'popup-price-label\' style=\'color:#2980b9\'>Detalle</div>"+' +
    '"<div class=\'popup-price-value\'>$"+detalle+"</div></div>"+' +
    '"<div class=\'popup-price\' style=\'background:#e8f5e9\'>"+' +
    '"<div class=\'popup-price-label\' style=\'color:#27ae60\'>Mayor</div>"+' +
    '"<div class=\'popup-price-value\'>$"+mayor+"</div></div></div>"+' +
    '"<div class=\'popup-stock\' style=\'background:"+stockBg+";color:"+stockColor+"\'>"+stockText+"</div>"+' +
    '"<div class=\'popup-barcode-num\'>"+bc+"</div>"+' +
    '"<button class=\'popup-btn\' onclick=\'addToVentas("+rowNum+")\' style=\'background:#1a3a2a\'>"+' +
    '"🛒 Agregar a Ventas / 판매추가</button>"+' +
    '"</div></div>";' +
    'document.getElementById("popupContainer").innerHTML=popup;' +
    'google.script.run.상품행이동(rowNum);}' +
    'function closePopup(){document.getElementById("popupContainer").innerHTML="";}' +
    'function addToVentas(rowNum){' +
    'var item=currentFiltered.find(function(i){return Number(i[5])===rowNum;});' +
    'if(!item)return;' +
    'closePopup();' +
    'google.script.run.withSuccessHandler(function(){})' +
    '.검색에서판매추가(item[0],item[1],item[2],item[3],rowNum);}' +
    'function clearAll(){' +
    'document.getElementById("searchInput").value="";' +
    'document.getElementById("clearBtn").style.display="none";' +
    'document.getElementById("statusLabel").textContent="✅ "+allData.length+"개 준비완료";' +
    'currentFiltered=[];' +
    'document.getElementById("results").innerHTML=' +
    '"<div class=\'no-result\'>Ingrese texto para buscar<br>검색어를 입력하세요</div>";}' +
    '<\/script></body></html>';
}

function getInventarioData() {
  // 캐시에서 먼저 확인
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('inventarioUIData');
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    result.push([
      String(data[i][0]).trim(),
      String(data[i][1]).trim(),
      String(data[i][2]).trim(),
      data[i][3] || 0,
      data[i][4] || 0,
      i + 2,
      Number(data[i][6]) || 0,
      ""
    ]);
  }

  // 캐시에 저장 (3분)
  try {
    CacheService.getScriptCache().put(
      'inventarioUIData', 
      JSON.stringify(result), 
      180
    );
  } catch(e) {}

  return result;
}

function 상품행이동(rowNum) {
  try {
    if (!rowNum || isNaN(rowNum)) return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
    if (!sheet) return;
    ss.setActiveSheet(sheet);

    // 기존 노란색 행 원래 색상으로 복원
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      for (var i = 2; i <= lastRow; i++) {
        var bg = i % 2 === 0 ? "#ffffff" : "#f0f7f0";
        sheet.getRange(i, 1, 1, 7).setBackground(bg);
      }
    }

    // 선택된 행만 노란색
    sheet.getRange(Number(rowNum), 1, 1, 7).setBackground("#FFE066");
    sheet.setActiveRange(sheet.getRange(Number(rowNum), 1));
  } catch(e) {
    console.log("상품행이동 오류: " + e.toString());
  }
}

function 검색에서판매추가(barcode, sku, producto, precio, rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 노란색 행 색상 제거
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (invSheet && rowNum) {
    var bg = Number(rowNum) % 2 === 0 ? "#ffffff" : "#f0f7f0";
    invSheet.getRange(Number(rowNum), 1, 1, 7).setBackground(bg);
  }

  // Ventas 시트로 이동 후 기록
  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (!ventasSheet) return;
  var lastRow = ventasSheet.getLastRow() + 1;
  if (lastRow < 2) lastRow = 2;
  ventasSheet.getRange(lastRow, 1, 1, 10).setValues([[
    new Date(), "", "", barcode, sku, producto, "", precio, "", ""
  ]]);
  ventasSheet.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm");
  ss.setActiveSheet(ventasSheet);
  ventasSheet.setActiveRange(ventasSheet.getRange(lastRow, 7));
}
// ========================
// 영수증
// ========================
function 출력BOLETA() { 영수증출력("BOLETA"); }
function 출력FACTURA() { 영수증출력("FACTURA"); }

function 영수증출력(tipo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (!ventasSheet || ventasSheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert("Ventas 시트에 판매 데이터가 없습니다!"); return; }
  var lastRow = ventasSheet.getLastRow();
  var data = ventasSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var items = data.filter(function(r) { return r[3] && r[6]; });
  if (items.length === 0) { SpreadsheetApp.getUi().alert("Ventas 시트에 유효한 판매 데이터가 없습니다!"); return; }
  영수증HTML생성(items, tipo || "BOLETA");
}

function 영수증HTML생성(items, tipo) {
  var TIENDA = "JK TRADING SPA";
  var DIRECCION = "AV FRANKLIN 543, SANTIAGO";
  var RUT = "77.530.277-1";
  var subtotal = 0;
  for (var i = 0; i < items.length; i++) subtotal += Number(items[i][8]) || 0;
  var neto = Math.round(subtotal / 1.19);
  var iva = subtotal - neto;
  var cliente = items[0][1] || "CLIENTE";
  var tipoPago = items[0][2] || "";
  var fechaStr = Utilities.formatDate(new Date(), "America/Santiago", "dd/MM/yyyy HH:mm");
  var itemsHtml = '';
  for (var j = 0; j < items.length; j++) {
    itemsHtml += '<tr><td style="padding:3px 2px;font-size:11px;border-bottom:1px dotted #ccc;">' + (items[j][5]||'') + '<br><span style="color:#888;font-size:10px;">' + (items[j][4]||'') + '</span></td><td style="padding:3px 2px;font-size:11px;text-align:center;border-bottom:1px dotted #ccc;">' + (items[j][6]||0) + '</td><td style="padding:3px 2px;font-size:11px;text-align:right;border-bottom:1px dotted #ccc;">$' + Number(items[j][7]||0).toLocaleString() + '</td><td style="padding:3px 2px;font-size:11px;text-align:right;border-bottom:1px dotted #ccc;">$' + Number(items[j][8]||0).toLocaleString() + '</td></tr>';
  }
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:"Courier New",monospace;background:#f5f5f5;padding:20px;display:flex;flex-direction:column;align-items:center;}.receipt{width:302px;background:white;padding:16px 12px;box-shadow:0 2px 10px rgba(0,0,0,0.15);border-radius:4px;}.header{text-align:center;border-bottom:2px dashed #333;padding-bottom:10px;margin-bottom:10px;}.tienda{font-size:16px;font-weight:bold;letter-spacing:1px;}.tipo-doc{font-size:14px;font-weight:bold;background:#1a2a3a;color:white;padding:4px 12px;border-radius:4px;display:inline-block;margin:6px 0;letter-spacing:2px;}.info{font-size:10px;color:#555;line-height:1.5;}.client-info{font-size:11px;border-bottom:1px dashed #333;padding-bottom:8px;margin-bottom:8px;line-height:1.6;}table{width:100%;border-collapse:collapse;}thead tr th{font-size:10px;text-align:left;padding:3px 2px;border-bottom:2px solid #333;border-top:1px solid #333;}thead tr th:nth-child(2){text-align:center;}thead tr th:nth-child(3),thead tr th:nth-child(4){text-align:right;}.totals{margin-top:8px;border-top:2px dashed #333;padding-top:8px;}.total-row{display:flex;justify-content:space-between;font-size:11px;padding:2px 0;}.total-row.main{font-size:14px;font-weight:bold;border-top:1px solid #333;margin-top:4px;padding-top:4px;}.total-row.iva{color:#555;font-size:11px;}.footer{text-align:center;margin-top:12px;padding-top:10px;border-top:2px dashed #333;font-size:10px;color:#888;line-height:1.8;}.btn-group{display:flex;gap:8px;margin-top:16px;width:302px;}.btn{flex:1;padding:10px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:bold;}.btn-print{background:#1a2a3a;color:white;}.btn-pdf{background:#1a3a2a;color:white;}.btn-close{background:#e74c3c;color:white;}@media print{body{background:white;padding:0;}.receipt{box-shadow:none;}.btn-group{display:none;}}</style></head><body>' +
    '<div class="receipt"><div class="header"><div class="tienda">' + TIENDA + '</div><div class="info">' + DIRECCION + '</div><div class="info">RUT: ' + RUT + '</div><div style="margin-top:6px;"><span class="tipo-doc">' + tipo + '</span></div><div class="info" style="margin-top:4px;">' + fechaStr + '</div></div>' +
    '<div class="client-info"><div><strong>Cliente:</strong> ' + cliente + '</div><div><strong>Forma de pago:</strong> ' + tipoPago + '</div></div>' +
    '<table><thead><tr><th>Producto</th><th>Cant</th><th>P.Unit</th><th>Total</th></tr></thead><tbody>' + itemsHtml + '</tbody></table>' +
    '<div class="totals"><div class="total-row"><span>Total:</span><span><strong>$' + subtotal.toLocaleString() + '</strong></span></div><div class="total-row iva"><span>IVA 19%:</span><span>$' + iva.toLocaleString() + '</span></div><div class="total-row iva"><span>Neto:</span><span>$' + neto.toLocaleString() + '</span></div><div class="total-row main"><span>NETO + IVA:</span><span>$' + subtotal.toLocaleString() + '</span></div></div>' +
    '<div class="footer"><div>¡Gracias por su compra!</div><div>Vuelva pronto</div></div></div>' +
    '<div class="btn-group"><button class="btn btn-print" onclick="window.print()">🖨️ Imprimir</button><button class="btn btn-pdf" onclick="window.print()">📄 Guardar PDF</button><button class="btn btn-close" onclick="google.script.host.close()">✖ Cerrar</button></div>' +
    '</body></html>'
  ).setWidth(380).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, tipo + ' - JK TRADING SPA');
}

// ========================
// 일일마감 보고서
// ========================
function 일일마감보고서생성() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var registroSheet = ss.getSheetByName(NOMBRE_REGISTRO);
  if (!registroSheet || registroSheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert("판매기록이 없습니다!"); return; }
  var today = new Date();
  var todayStr = Utilities.formatDate(today, "America/Santiago", "yyyy-MM-dd");
  var lastRow = registroSheet.getLastRow();
  var allData = registroSheet.getRange(2, 1, lastRow-1, 10).getValues();
  var todayData = allData.filter(function(row) {
    if (!row[0]) return false;
    return Utilities.formatDate(new Date(row[0]), "America/Santiago", "yyyy-MM-dd") === todayStr;
  });
  if (todayData.length === 0) { SpreadsheetApp.getUi().alert("오늘 판매 데이터가 없습니다!\n날짜: " + todayStr); return; }
  var reportName = "마감보고서_" + todayStr;
  var reportSheet = ss.getSheetByName(reportName);
  if (!reportSheet) reportSheet = ss.insertSheet(reportName);
  else { reportSheet.clearContents(); reportSheet.clearFormats(); }
  var writeRow = 1;
  reportSheet.getRange(writeRow,1,1,5).setValues([["📊 일일마감보고서 - "+todayStr,"","","",""]]).setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
  reportSheet.getRange(writeRow,1,1,5).merge(); reportSheet.setRowHeight(writeRow,40); writeRow++;
  reportSheet.getRange(writeRow,1,1,5).setValues([["JK TRADING SPA | AV FRANKLIN 543 | RUT: 77.530.277-1","","","",""]]).setBackground("#2c3e50").setFontColor("#FFFFFF").setFontSize(11).setHorizontalAlignment("center");
  reportSheet.getRange(writeRow,1,1,5).merge(); writeRow+=2;
  var totalVentas=0,totalItems=0;
  for(var i=0;i<todayData.length;i++){totalVentas+=Number(todayData[i][8])||0;totalItems+=Number(todayData[i][6])||0;}
  var IVA=Math.round(totalVentas-(totalVentas/1.19));
  var neto=totalVentas-IVA;
  reportSheet.getRange(writeRow,1,1,5).setValues([["💰 매출 요약","","","",""]]).setBackground("#1a3a2a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  reportSheet.getRange(writeRow,1,1,5).merge(); writeRow++;
  var summaryData=[["판매건수",todayData.length+"건","","판매수량",totalItems+"개"],["총 매출","$"+totalVentas.toLocaleString(),"","IVA 19%","$"+IVA.toLocaleString()],["Neto","$"+neto.toLocaleString(),"","Neto + IVA","$"+totalVentas.toLocaleString()]];
  reportSheet.getRange(writeRow,1,summaryData.length,5).setValues(summaryData);
  for(var r=0;r<summaryData.length;r++){reportSheet.getRange(writeRow+r,1,1,5).setBackground(r%2===0?"#f0f7f0":"#ffffff");reportSheet.getRange(writeRow+r,1).setFontWeight("bold").setFontColor("#1a3a2a");reportSheet.getRange(writeRow+r,4).setFontWeight("bold").setFontColor("#1a3a2a");}
  writeRow+=summaryData.length+2;
  var invSheet=ss.getSheetByName(NOMBRE_INVENTARIO);
  reportSheet.setColumnWidth(1,160);reportSheet.setColumnWidth(2,120);reportSheet.setColumnWidth(3,200);reportSheet.setColumnWidth(4,100);reportSheet.setColumnWidth(5,120);
  ss.setActiveSheet(reportSheet);
  SpreadsheetApp.getUi().alert("✅ 일일마감 보고서 생성완료!\n시트명: " + reportName);
}

function 영업종료자동보고서() {
  var hour = new Date().getHours();
  var min = new Date().getMinutes();
  if (hour === 17 && min >= 30) 일일마감보고서생성();
}

// ========================
// 재고실사
// ========================
function 재고실사시트생성() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("재고실사");
  var ui = SpreadsheetApp.getUi();
  if (sheet) {
    var response = ui.alert("기존 재고실사 시트가 있습니다!","새로 만들까요?\n\n확인 = 새로만들기\n취소 = 기존것 열기",ui.ButtonSet.OK_CANCEL);
    if (response === ui.Button.CANCEL) { ss.setActiveSheet(sheet); return; }
    sheet.clearContents(); sheet.clearFormats();
  } else {
    sheet = ss.insertSheet("재고실사");
  }
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) { ui.alert("Inventario 시트를 찾을 수 없습니다!"); return; }
  var lastRow = invSheet.getLastRow();
  var invData = invSheet.getRange(2,1,lastRow-1,7).getValues();
  var headers = ["바코드","SKU","상품명","시스템재고","실제재고","차이","상태","비고"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1,35); sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,140);sheet.setColumnWidth(2,120);sheet.setColumnWidth(3,220);
  sheet.setColumnWidth(4,90);sheet.setColumnWidth(5,90);sheet.setColumnWidth(6,80);
  sheet.setColumnWidth(7,100);sheet.setColumnWidth(8,150);
  var rows=[];
  for(var i=0;i<invData.length;i++){
    if(!invData[i][0]&&!invData[i][1])continue;
    rows.push([String(invData[i][0]).trim(),String(invData[i][1]).trim(),String(invData[i][2]).trim(),Number(invData[i][6])||0,"","","",""]);
  }
  if(rows.length>0){
    sheet.getRange(2,1,rows.length,8).setValues(rows);
    for(var r=0;r<rows.length;r++){
      sheet.getRange(r+2,6).setFormula('=IF(E'+(r+2)+'="","",E'+(r+2)+'-D'+(r+2)+')');
      sheet.getRange(r+2,1,1,8).setBackground(r%2===0?"#f0f4f7":"#ffffff");
    }
    sheet.getRange(2,4,rows.length,1).setBackground("#f5f5f5").setFontColor("#888888");
    sheet.getRange(2,5,rows.length,1).setBackground("#fff9e6").setBorder(true,true,true,true,true,true,"#f0c040",SpreadsheetApp.BorderStyle.SOLID);
  }
  sheet.getRange(1,10).setValue("실사일자: "+Utilities.formatDate(new Date(),"America/Santiago","yyyy-MM-dd HH:mm")).setFontWeight("bold").setFontColor("#1a2a3a");
  ss.setActiveSheet(sheet);
  ui.alert("✅ 재고실사 시트 생성완료!\n"+rows.length+"개 상품\n\n① E열에 실제재고 입력\n② 완료 버튼 클릭");
}

function 재고실사onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();
  if (row < 2) return;
  if (col === 1) {
    var barcode = String(range.getValue()).replace(/[^0-9]/g,'');
    if (!barcode) return;
    range.setValue(barcode);
    var invData = getInvData();
    if (!invData) return;
    for (var i = 0; i < invData.length; i++) {
      if (String(invData[i][0]).replace(/[^0-9]/g,'') === barcode) {
        sheet.getRange(row,2).setValue(invData[i][1]);
        sheet.getRange(row,3).setValue(invData[i][2]);
        sheet.getRange(row,4).setValue(Number(invData[i][6])||0);
        sheet.setActiveRange(sheet.getRange(row,5));
        break;
      }
    }
  }
  if (col === 5) {
    var actual = Number(range.getValue());
    var system = Number(sheet.getRange(row,4).getValue());
    var diff = actual - system;
    if (!isNaN(diff)) {
      var status = diff===0?"✅ 일치":diff>0?"▲ +"+diff+" 초과":"▼ "+diff+" 부족";
      var statusColor = diff===0?"#1a3a2a":diff>0?"#2980b9":"#e74c3c";
      var bg = diff===0?"#e8f5e9":diff>0?"#eaf0fb":"#fadbd8";
      sheet.getRange(row,6).setValue(diff);
      sheet.getRange(row,7).setValue(status).setFontColor(statusColor).setFontWeight("bold").setBackground(bg);
      sheet.getRange(row,5).setBackground(bg);
    }
  }
}

function 재고실사결과반영() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("재고실사");
  if (!sheet) { SpreadsheetApp.getUi().alert("재고실사 시트가 없습니다!"); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("데이터가 없습니다!"); return; }
  var data = sheet.getRange(2,1,lastRow-1,7).getValues();
  var diffRows = data.filter(function(r){return r[4]!==""&&r[5]!==""&&r[5]!==0;});
  if (diffRows.length === 0) { SpreadsheetApp.getUi().alert("✅ 차이나는 상품이 없습니다!"); return; }
  var ui = SpreadsheetApp.getUi();
  if (ui.alert("재고실사 결과 반영","차이나는 상품 "+diffRows.length+"개를 반영할까요?",ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invData = invSheet.getRange(2,1,invSheet.getLastRow()-1,7).getValues();
  var invMap = {};
  for (var i = 0; i < invData.length; i++) invMap[String(invData[i][0]).trim()] = i;
  var count = 0;
  for (var j = 0; j < diffRows.length; j++) {
    var barcode = String(diffRows[j][0]).trim();
    var actual = Number(diffRows[j][4]);
    if (!barcode || isNaN(actual)) continue;
    if (invMap.hasOwnProperty(barcode)) {
      var idx = invMap[barcode];
      var oldStock = Number(invData[idx][6]) || 0;
      invData[idx][6] = actual;
      count++;
    }
  }
  invSheet.getRange(2,1,invData.length,7).setValues(invData);
  invCacheClear();
  통합재고동기화();
  SpreadsheetApp.getUi().alert("✅ 재고실사 완료!\n"+count+"개 상품 업데이트!");
}

function 재고실사차이표시() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("재고실사");
  if (!sheet) { SpreadsheetApp.getUi().alert("재고실사 시트가 없습니다!"); return; }
  ss.setActiveSheet(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.showRows(2, lastRow-1);
  var data = sheet.getRange(2,1,lastRow-1,7).getValues();
  var hideStart=-1; var diffCount=0;
  for (var i = 0; i < data.length; i++) {
    var actual=data[i][4]; var diff=data[i][5]; var rowNum=i+2;
    var shouldHide=actual!==""&&(diff===0||diff==="");
    if(shouldHide){if(hideStart===-1)hideStart=rowNum;}
    else{if(hideStart!==-1){sheet.hideRows(hideStart,rowNum-hideStart);hideStart=-1;}if(actual!=="")diffCount++;}
  }
  if(hideStart!==-1)sheet.hideRows(hideStart,lastRow-hideStart+1);
  SpreadsheetApp.getUi().alert("차이나는 상품: "+diffCount+"개 표시중");
}

function 재고실사전체보기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("재고실사");
  if (!sheet) return;
  ss.setActiveSheet(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.showRows(2, lastRow-1);
}

function aplicarInventario() { 재고실사결과반영(); }

// ========================
// 사진 관련
// ========================
function 사진열추가_Inventario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) { SpreadsheetApp.getUi().alert("Inventario 시트를 찾을 수 없습니다!"); return; }
  var photoMap = {};
  try {
    var 통합ss = SpreadsheetApp.openById(ID_통합재고);
    var urlSheet = 통합ss.getSheetByName("사진URL");
    if (!urlSheet) { SpreadsheetApp.getUi().alert("통합재고 파일에 사진URL 시트가 없습니다!"); return; }
    var urlData = urlSheet.getDataRange().getValues();
    for (var i = 1; i < urlData.length; i++) {
      var bc = String(urlData[i][0]).trim();
      var rawUrl = String(urlData[i][2]).trim().replace(/"/g,'').replace(/\s/g,'');
      var match = rawUrl.match(/id=([a-zA-Z0-9_-]+)/);
      var url = match ? "https://lh3.googleusercontent.com/d/" + match[1] : rawUrl;
      if (bc && url) photoMap[bc] = url;
    }
  } catch(e) { SpreadsheetApp.getUi().alert("통합재고 파일 접근 오류!\n" + e.message); return; }
  invSheet.getRange(1,8).setValue("Foto").setBackground("#1a3a2a").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  invSheet.setColumnWidth(8, 80);
  var lastRow = invSheet.getLastRow();
  if (lastRow < 2) return;
  var barcodes = invSheet.getRange(2, 1, lastRow-1, 1).getValues();
  var formulas = []; var count = 0; var noPhoto = 0;
  for (var j = 0; j < barcodes.length; j++) {
    var barcode = String(barcodes[j][0]).trim();
    var photoUrl = photoMap[barcode];
    if (photoUrl) { formulas.push(["=IMAGE(\"" + photoUrl + "\")"]); count++; }
    else { formulas.push(["📷"]); noPhoto++; }
  }
  invSheet.getRange(2, 8, barcodes.length, 1).setValues(formulas).setHorizontalAlignment("center").setVerticalAlignment("middle");
  for (var r = 2; r <= lastRow; r++) invSheet.setRowHeight(r, 55);
  invCacheClear();
  SpreadsheetApp.getUi().alert("✅ 사진 연결 완료!\n사진 있음: "+count+"개\n사진 없음: "+noPhoto+"개");
}

function 사진열토글() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) return;
  var isHidden = sheet.isColumnHiddenByUser(8);
  if (isHidden) { sheet.showColumns(8); SpreadsheetApp.getUi().alert("📷 사진열 표시됨!"); }
  else { sheet.hideColumns(8); SpreadsheetApp.getUi().alert("📷 사진열 숨겨짐!"); }
}

function H열초기화() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 8, lastRow-1, 1).clearContent();
  sheet.getRange(2, 8, lastRow-1, 1).clearFormat();
  for (var r = 2; r <= lastRow; r++) sheet.setRowHeight(r, 21);
  invCacheClear();
  SpreadsheetApp.getUi().alert("✅ H열 초기화 완료!");
}

function 시트숨기기관리() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 숨길시트 = [
    "매출분석",
    "COTIZACION",
    "사진URL",
    "재고실사",
    "재고수동조정"
  ];
  var sheets = ss.getSheets();
  var count = 0;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (숨길시트.indexOf(name) !== -1 || 
        name.indexOf("마감보고서_") !== -1) {
      sheets[i].hideSheet();
      count++;
    }
  }
  SpreadsheetApp.getUi().alert("✅ " + count + "개 시트 숨김 완료!");
}

function 시트보이기관리() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 보일시트 = [
    "매출분석",
    "COTIZACION",
    "사진URL",
    "재고실사",
    "재고수동조정"
  ];
  var sheets = ss.getSheets();
  var count = 0;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (보일시트.indexOf(name) !== -1 || 
        name.indexOf("마감보고서_") !== -1) {
      sheets[i].showSheet();
      count++;
    }
  }
  SpreadsheetApp.getUi().alert("✅ " + count + "개 시트 표시 완료!");
}

function 바코드직접생성() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) { SpreadsheetApp.getUi().alert("Inventario 시트를 찾을 수 없습니다!"); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow-1, 2).getValues();
  var maxBarcode = 2000000000000;
  for (var i = 0; i < data.length; i++) {
    var bc = Number(String(data[i][0]).replace(/[^0-9]/g,''));
    if (bc > maxBarcode) maxBarcode = bc;
  }
  var count = 0;
  for (var i = 0; i < data.length; i++)
    if (!data[i][0] && data[i][1]) { maxBarcode++; data[i][0] = maxBarcode; count++; }
  if (count === 0) { SpreadsheetApp.getUi().alert("바코드가 없는 SKU가 없습니다!"); return; }
  sheet.getRange(2, 1, data.length, 1).setValues(data.map(function(r){return [r[0]];}));
  invCacheClear();
  SpreadsheetApp.getUi().alert("✅ " + count + "개 바코드 생성완료!");
}

function 부족재고표시_자동() {
  var invSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_INVENTARIO);
  if (!invSheet) return;
  var lastRow = invSheet.getLastRow();
  if (lastRow >= 2) invSheet.getRange(2, 9, lastRow-1, 2).clearContent();
}

function 재고빈셀0처리() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 7, lastRow-1, 1).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "" || data[i][0] === null) {
      data[i][0] = 0;
      count++;
    }
  }
  sheet.getRange(2, 7, lastRow-1, 1).setValues(data);
  SpreadsheetApp.getUi().alert("✅ " + count + "개 빈셀을 0으로 처리완료!");
}

function G열색상초기화() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  if (!sheet) { 
    SpreadsheetApp.getUi().alert("Inventario 시트를 찾을 수 없습니다!"); 
    return; 
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 7, lastRow - 1, 1)
    .setFontColor("#000000")
    .setNumberFormat("0")
    .setHorizontalAlignment("center");
  SpreadsheetApp.getUi().alert("✅ G열 색상 초기화 완료!\n" + (lastRow - 1) + "개 셀 적용됨");
}

function 시트권한일괄설정() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 특정인 이메일
  var 특정인이메일 = "dmsgh115@gmail.com";
  
  // 소유자 + 특정인만 편집 가능한 시트
  var 보호시트 = [
    "매출분석",
    "COTIZACION",
    "사진URL",
    "재고실사",
    "재고수동조정"
  ];
  
  // 마감보고서_ 로 시작하는 시트도 자동 추가
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().indexOf("마감보고서_") !== -1) {
      보호시트.push(allSheets[i].getName());
    }
  }
  
  var count = 0;
  for (var j = 0; j < 보호시트.length; j++) {
    var sheet = ss.getSheetByName(보호시트[j]);
    if (!sheet) continue;
    
    // 기존 보호 제거
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var p = 0; p < protections.length; p++) {
      protections[p].remove();
    }
    
    // 새 보호 설정
    var protection = sheet.protect();
    protection.setDescription(보호시트[j] + " 보호");
    
    // 특정인 추가
    protection.addEditor(특정인이메일);
    
    // 소유자와 특정인 외 모두 제거
    var editors = protection.getEditors();
    for (var e = 0; e < editors.length; e++) {
      var email = editors[e].getEmail();
      if (email !== 특정인이메일 &&
          email !== Session.getActiveUser().getEmail()) {
        protection.removeEditor(editors[e]);
      }
    }
    
    // 시트 숨기기
    sheet.hideSheet();
    count++;
  }
  
  SpreadsheetApp.getUi().alert(
    "✅ 시트 보호 완료!\n" + count + "개 시트 설정됨\n\n" +
    "보호 + 숨김 적용:\n" + 보호시트.join("\n")
  );
}

// ========================
// Eroom 발주서 → NAOS 판매시트 전송
// ========================
const EROOM_ID = "1hroizqPQjd40EkAEa80WMtU5aLFYPwuIINv9QtykYaM";

function sendEroomOrderToNAOS() {
  var ui = SpreadsheetApp.getUi();
  
  // Eroom 발주서 시트 가져오기
  var eroomSS = SpreadsheetApp.openById(EROOM_ID);
  var sheets = eroomSS.getSheets();
  
  // 발주서_ 로 시작하는 시트 찾기
  var orderSheet = null;
  for (var i = sheets.length - 1; i >= 0; i--) {
    if (sheets[i].getName().indexOf("발주서_") === 0) {
      orderSheet = sheets[i];
      break;
    }
  }
  
  if (!orderSheet) {
    ui.alert("⚠️ Eroom에서 발주서 시트를 찾을 수 없어요!\n발주서_날짜 형식의 시트가 있는지 확인해주세요.");
    return;
  }
  
  var sheetName = orderSheet.getName();
  var orderData = orderSheet.getDataRange().getValues();
  
  if (orderData.length < 2) {
    ui.alert("⚠️ 발주서 데이터가 없어요!");
    return;
  }
  
  // 헤더에서 컬럼 위치 찾기
  var headers = orderData[0];
  var skuCol   = -1, qtyCol = -1;
  for (var h = 0; h < headers.length; h++) {
    var hv = String(headers[h]).trim();
    if (hv === "SKU" || hv === "Codigo SKU") skuCol = h;
    if (hv === "발주수량" || hv === "Cantidad") qtyCol = h;
  }
  
  // 헤더에서 못 찾으면 고정 위치 사용 (C=2, G=6)
  if (skuCol === -1) skuCol = 2;
  if (qtyCol === -1) qtyCol = 6;
  
  // 유효한 행 수집
  var validRows = [];
  for (var i = 1; i < orderData.length; i++) {
    var sku = String(orderData[i][skuCol] || "").trim();
    var qty = Number(orderData[i][qtyCol]);
    if (sku && qty > 0) {
      validRows.push({ sku: sku, qty: qty });
    }
  }
  
  if (validRows.length === 0) {
    ui.alert("⚠️ 전송할 데이터가 없어요!\n발주수량이 입력된 행이 있는지 확인해주세요.");
    return;
  }
  
  // NAOS 재고에서 SKU → 바코드/상품명/단가 조회
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invData = invSheet.getDataRange().getValues();
  
  var naosMap = {};
  for (var j = 1; j < invData.length; j++) {
    var s = String(invData[j][1] || "").trim();
    if (s) {
      naosMap[s] = {
        barcode:     String(invData[j][0] || "").trim(),
        nombre:      String(invData[j][2] || "").trim(),
        priceDetalle: Number(invData[j][3]) || 0,
        priceMayor:   Number(invData[j][4]) || 0
      };
    }
  }
  
  // NAOS Ventas 시트에 기록
  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (!ventasSheet) {
    ui.alert("⚠️ NAOS Ventas 시트를 찾을 수 없어요!");
    return;
  }
  
  var today = new Date();
  var notFound = [];
  var rows = [];
  
  for (var k = 0; k < validRows.length; k++) {
    var sku  = validRows[k].sku;
    var qty  = validRows[k].qty;
    var info = naosMap[sku];
    
    if (!info) { notFound.push(sku); continue; }
    
    rows.push([
      today,           // A: 날짜
      "",              // B: 거래처 (수동 선택)
      "",              // C: 결제방법
      info.barcode,    // D: 바코드
      sku,             // E: SKU
      info.nombre,     // F: 상품명
      qty,             // G: 수량
      info.priceDetalle, // H: 단가 (소매가 기본)
      qty * info.priceDetalle, // I: 금액
      ""               // J: 합계
    ]);
  }
  
  if (rows.length === 0) {
    ui.alert("⚠️ NAOS 재고에서 일치하는 SKU가 없어요!\n없는 SKU: " + notFound.join(", "));
    return;
  }
  
  var lastRow = ventasSheet.getLastRow() + 1;
  if (lastRow < 2) lastRow = 2;
  
  ventasSheet.getRange(lastRow, 1, rows.length, 10).setValues(rows);
  ventasSheet.getRange(lastRow, 1, rows.length, 1).setNumberFormat("yyyy-MM-dd HH:mm");
  ventasSheet.getRange(lastRow, 8, rows.length, 2).setNumberFormat("#,##0");
  
  // 거래처 드롭다운 설정
  var clientSheet = ss.getSheetByName(NOMBRE_CLIENTES);
  if (clientSheet) {
    var clientData = clientSheet.getDataRange().getValues();
    var clientNames = clientData.map(function(r) { return r[0]; }).filter(Boolean);
    if (clientNames.length > 0) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(clientNames, true)
        .build();
      ventasSheet.getRange(lastRow, 2, rows.length, 1).setDataValidation(rule);
    }
  }
  
  // 합계 업데이트
  actualizarTotal(ventasSheet);
  
  ss.setActiveSheet(ventasSheet);
  ventasSheet.setActiveRange(ventasSheet.getRange(lastRow, 2)); // 거래처 선택으로 커서
  
  var msg = "✅ 발주서 → NAOS 판매시트 전송완료!\n" + rows.length + "개 상품\n발주서: " + sheetName;
  if (notFound.length > 0) msg += "\n\n❌ SKU 없음: " + notFound.join(", ");
  ui.alert(msg);
}

// ========================
// ML 판매 자동화 + 묶음 연동
// ========================

// ML API 설정 (PropertiesService에 저장된 토큰 사용)
// ========================
// ML 토큰 관리 (시트 기반)
// ========================

function getMLTokenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ML토큰설정");
  if (sheet) return sheet;

  // 시트 없으면 자동 생성
  sheet = ss.insertSheet("ML토큰설정");
  var headers = ["항목", "LUCAS 계정", "NAOS 계정"];
  sheet.getRange(1, 1, 1, 3).setValues([headers])
    .setBackground("#1a2a3a").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(2, 1, 6, 1).setValues([
    ["CLIENT_ID"],
    ["CLIENT_SECRET"],
    ["ACCESS_TOKEN"],
    ["REFRESH_TOKEN"],
    ["SELLER_ID"],
    ["토큰만료시간"]
  ]).setFontWeight("bold");

  // 기본값 입력
  sheet.getRange(2, 2).setValue("30482610593775"); // Lucas CLIENT_ID
  sheet.getRange(5, 2).setValue("323449286");       // Lucas SELLER_ID

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 320);

  SpreadsheetApp.getUi().alert(
    "✅ ML토큰설정 시트 생성!\n\n" +
    "B열(LUCAS), C열(NAOS)에\n" +
    "ACCESS_TOKEN과 REFRESH_TOKEN을 입력해주세요."
  );
  return sheet;
}

function getMLToken(cuenta) {
  var sheet = getMLTokenSheet();
  var col   = (cuenta === "NAOS") ? 3 : 2;

  var accessToken  = String(sheet.getRange(4, col).getValue() || "").trim();
  var refreshToken = String(sheet.getRange(5, col).getValue() || "").trim();
  var expiry       = sheet.getRange(7, col).getValue();
  var now          = new Date().getTime();

  // 토큰 없으면 빈값 반환
  if (!accessToken && !refreshToken) return "";

  // 만료됐거나 만료 10분 전이면 자동 갱신
  var needRefresh = false;
  if (!expiry) {
    needRefresh = true; // 만료시간 없으면 갱신 시도
  } else if (now > Number(expiry) - 600000) {
    needRefresh = true; // 10분 전부터 갱신
  }

  if (needRefresh && refreshToken) {
    console.log(cuenta + " 토큰 갱신 시도...");
    var refreshed = refreshMLToken(cuenta);
    if (refreshed) {
      // 갱신 성공 → 새 토큰 반환
      return String(sheet.getRange(4, col).getValue() || "").trim();
    } else {
      console.log(cuenta + " 토큰 갱신 실패 — 기존 토큰으로 시도");
    }
  }

  return accessToken;
}

function manualRefreshMLToken() {
  var ui = SpreadsheetApp.getUi();

  // LUCAS 갱신
  var lucasResult = refreshMLToken("LUCAS");

  var msg = "ML 토큰 갱신 결과:\n\n";
  msg += "LUCAS: " + (lucasResult ? "✅ 갱신완료" : "❌ 실패") + "\n";

  if (!lucasResult) {
    msg += "\n갱신 실패 원인:\n";
    msg += "• REFRESH_TOKEN이 만료됐을 수 있어요\n";
    msg += "• 이 경우 ML에서 새로 로그인 후\n";
    msg += "  새 토큰을 ML토큰설정 시트에 입력해야 해요\n\n";
    msg += "새 토큰 발급 URL:\n";
    msg += "https://auth.mercadolibre.cl/authorization?response_type=code";
    msg += "&client_id=30482610593775&redirect_uri=https://www.google.com";
  }

  ui.alert(msg);
}

function refreshMLToken(cuenta) {
  var sheet = getMLTokenSheet();
  var col   = (cuenta === "NAOS") ? 3 : 2;

  var clientId     = String(sheet.getRange(2, col).getValue() || "").trim();
  var clientSecret = String(sheet.getRange(3, col).getValue() || "").trim();
  var refreshToken = String(sheet.getRange(5, col).getValue() || "").trim();

  if (!clientId || !clientSecret || !refreshToken) {
    console.log(cuenta + " 갱신 실패: 설정값 없음 " +
      "ID:" + clientId.substring(0,10) +
      " SECRET:" + clientSecret.substring(0,5) +
      " REFRESH:" + refreshToken.substring(0,10));
    return false;
  }

  try {
    var response = UrlFetchApp.fetch("https://api.mercadolibre.com/oauth/token", {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: "grant_type=refresh_token" +
               "&client_id="     + clientId +
               "&client_secret=" + clientSecret +
               "&refresh_token=" + refreshToken,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.log(cuenta + " 갱신 실패: " + response.getContentText());
      return false;
    }

    var result     = JSON.parse(response.getContentText());
    var expiryTime = new Date().getTime() + (result.expires_in * 1000);

    sheet.getRange(4, col).setValue(result.access_token);
    sheet.getRange(5, col).setValue(result.refresh_token);
    sheet.getRange(7, col).setValue(expiryTime);

    console.log("✅ " + cuenta + " 토큰 자동 갱신 완료");
    return true;

  } catch(err) {
    console.log(cuenta + " 갱신 오류: " + err.message);
    return false;
  }
}

function mlTokenSetupHelper() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getMLTokenSheet();

  // 현재 입력된 값 확인
  var lucasAccess = sheet.getRange(3, 2).getValue();
  var naosAccess  = sheet.getRange(3, 3).getValue();

  var msg = "📋 ML 토큰 설정 현황\n\n";
  msg += "LUCAS: " + (lucasAccess ? "✅ 입력됨" : "❌ 없음") + "\n";
  msg += "NAOS:  " + (naosAccess  ? "✅ 입력됨" : "❌ 없음") + "\n\n";
  msg += "ML토큰설정 시트에서 직접 입력해주세요.\n\n";
  msg += "필요한 항목:\n";
  msg += "• CLIENT_ID\n• CLIENT_SECRET\n";
  msg += "• ACCESS_TOKEN\n• REFRESH_TOKEN\n• SELLER_ID";

  ui.alert(msg);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}

// ── 묶음설정 시트 확인/생성 ──────────────────────────────
function 묶음설정시트확인() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("묶음설정");
  if (sheet) return sheet;

  sheet = ss.insertSheet("묶음설정");
  var headers = ["SKU", "상품명", "ML_ID", "묶음수량", "계정(LUCAS/NAOS)", "비고"];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground("#1a2a3a").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
  [120, 200, 160, 90, 130, 150].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
  SpreadsheetApp.getUi().alert(
    "✅ 묶음설정 시트 생성완료!\n\n" +
    "사용법:\n" +
    "A열: SKU (NAOS재고 SKU와 동일)\n" +
    "B열: 상품명\n" +
    "C열: ML 상품 ID (예: MLC12345678)\n" +
    "D열: 묶음수량 (예: 10 → ML 1개 판매 = 재고 10개 차감)\n" +
    "E열: 계정 (LUCAS 또는 NAOS)\n\n" +
    "묶음수량이 1이면 일반 상품과 동일하게 처리됩니다."
  );
  return sheet;
}

// ── ML 판매내역 가져오기 + 재고 차감 ────────────────────
function fetchMLSalesAndDeduct() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 묶음설정 확인
  var bundleSheet = ss.getSheetByName("묶음설정");
  if (!bundleSheet || bundleSheet.getLastRow() < 2) {
    ui.alert("⚠️ 묶음설정 시트가 없거나 비어있어요!"); return;
  }
  var bundleData = bundleSheet.getRange(2,1,bundleSheet.getLastRow()-1,5).getValues();
  var bundleMap  = {};
  for (var i = 0; i < bundleData.length; i++) {
    var mlId   = String(bundleData[i][2]||"").trim();
    var sku    = String(bundleData[i][0]||"").trim();
    var bundle = Number(bundleData[i][3])||1;
    if (mlId && sku) bundleMap[mlId] = {sku:sku, bundle:bundle};
  }

  // 완료주문번호 (중복방지)
  var doneSheet = ss.getSheetByName("✅완료주문번호");
  if (!doneSheet) {
    doneSheet = ss.insertSheet("✅완료주문번호");
    doneSheet.getRange(1,1).setValue("주문번호")
      .setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold");
  }
  var doneIds = {};
  if (doneSheet.getLastRow() >= 2) {
    var doneData = doneSheet.getRange(2,1,doneSheet.getLastRow()-1,1).getValues();
    for (var d = 0; d < doneData.length; d++) {
      var did = String(doneData[d][0]).trim();
      if (did) doneIds[did] = true;
    }
  }

  // 토큰
  var token = getMLToken("LUCAS");
  if (!token) {
    ui.alert("❌ LUCAS 토큰 없음\n메뉴 → 토큰 수동 갱신을 실행해주세요.");
    return;
  }

  // 최근 30일 미처리 주문 조회
  var dateFrom = new Date();
  dateFrom.setDate(dateFrom.getDate() - 30);
  var dateFromStr = dateFrom.toISOString();

  var url = "https://api.mercadolibre.com/orders/search" +
            "?seller=323449286" +
            "&order.status=paid" +
            "&order.date_created.from=" + dateFromStr +
            "&sort=date_desc&limit=50";

  var response = UrlFetchApp.fetch(url, {
    headers: {"Authorization": "Bearer " + token},
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    ui.alert(
      "⚠️ ML API 오류(" + response.getResponseCode() + ")\n\n" +
      "오류내용: " + response.getContentText().substring(0,200) + "\n\n" +
      "URL: " + url
    );
    return;
  }

  var orders = JSON.parse(response.getContentText()).results || [];
  if (orders.length === 0) {
    ui.alert("✅ 새로운 ML 판매 없음");
    return;
  }

  // ML판매내역 시트 준비
  var salesSheet = ss.getSheetByName("ML판매내역");
  if (!salesSheet) {
    salesSheet = ss.insertSheet("ML판매내역");
  }

  // 헤더 확인 및 설정
  var hdrs = [
    "선택", "처리일시", "주문ID", "계정", "ML_ID", "SKU",
    "상품명", "ML/FLEX", "ML판매수량", "묶음수량",
    "실제차감수량", "ML단가", "금액", "처리결과", "URL"
  ];
  salesSheet.getRange(1,1,1,hdrs.length).setValues([hdrs])
    .setBackground("#1a3a2a").setFontColor("#FFFFFF").setFontWeight("bold");
  salesSheet.setFrozenRows(1);

  // 열 너비
  var colWidths = [40,120,160,60,120,100,200,60,70,70,80,80,90,100,200];
  colWidths.forEach(function(w,i){ salesSheet.setColumnWidth(i+1, w); });

  // 재고 시트
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invData  = invSheet.getRange(2,1,invSheet.getLastRow()-1,7).getValues();
  var invMap   = {};
  for (var im = 0; im < invData.length; im++) {
    var s = String(invData[im][1]||"").trim();
    if (s) invMap[s] = im;
  }

  var newDoneIds  = [];
  var logRows     = [];
  var deductCount = 0;
  var skipCount   = 0;
  var newCount    = 0;
  var notFound    = [];
  var notInBundle = [];

  for (var o = 0; o < orders.length; o++) {
    var order    = orders[o];
    var rawOrder = JSON.stringify(order);

    // 16자리 주문번호 정밀도 손실 방지
    var idMatch = rawOrder.match(/"id"\s*:\s*(\d{15,})/);
    var orderId = idMatch ? idMatch[1] : String(order.id);

    // 이미 처리된 주문 스킵
    if (doneIds[orderId]) { skipCount++; continue; }

    // ML/FLEX 구분 — shipping ID로 상세 조회
    var mlFlexType = "ML"; // 기본값
    var shippingId = order.shipping ? String(order.shipping.id||"") : "";
    if (shippingId) {
      try {
        var shipResp = UrlFetchApp.fetch(
          "https://api.mercadolibre.com/shipments/" + shippingId,
          {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
        );
        if (shipResp.getResponseCode() === 200) {
          var shipData    = JSON.parse(shipResp.getContentText());
          var logisticType= String(shipData.logistic_type||"").toLowerCase();
          if (logisticType === "self_service" ||
              logisticType.indexOf("fulfillment") !== -1 ||
              logisticType.indexOf("flex") !== -1) {
            mlFlexType = "FLEX";
          }
        }
      } catch(e) {
        console.log("shipping 조회 오류: " + e.message);
      }
    }

    // 주문일시
    var orderDate = order.date_created
      ? new Date(order.date_created) : new Date();

    var items = order.order_items || [];
    var hasItem = false;

    for (var oi = 0; oi < items.length; oi++) {
      var item     = items[oi];
      var mlItemId = item.item ? String(item.item.id) : "";
      var mlQty    = Number(item.quantity)||0;
      var nombre   = item.item ? (item.item.title||"") : "";
      var unitPrice= Number(item.unit_price)||0;

      if (!mlItemId || !mlQty) continue;

      // 묶음설정에 없는 상품
      if (!bundleMap[mlItemId]) {
        notInBundle.push(mlItemId + " (" + nombre.substring(0,20) + ")");
        continue;
      }

      var sku     = bundleMap[mlItemId].sku;
      var bundle  = bundleMap[mlItemId].bundle;
      var realQty = mlQty * bundle;
      var amount  = unitPrice * mlQty;
      var result  = "";

      // 재고 차감
      if (invMap.hasOwnProperty(sku)) {
        var idx      = invMap[sku];
        var curStock = Number(invData[idx][6])||0;
        var newStock = Math.max(0, curStock - realQty);
        invData[idx][6] = newStock;
        result = "✅ " + curStock + " → " + newStock;
        deductCount++;
      } else {
        result = "❌ SKU 없음";
        notFound.push(sku);
      }

      // ML 상품 URL
      var itemUrl = mlItemId
        ? 'https://articulo.mercadolibre.cl/MLC-' + mlItemId.replace('MLC','')
        : '';

      logRows.push([
        false,        // A: 체크박스 (선택)
        orderDate,    // B: 처리일시
        orderId,      // C: 주문ID
        "LUCAS",      // D: 계정
        mlItemId,     // E: ML_ID
        sku,          // F: SKU
        nombre,       // G: 상품명
        mlFlexType,   // H: ML/FLEX
        mlQty,        // I: ML판매수량
        bundle,       // J: 묶음수량
        realQty,      // K: 실제차감수량
        unitPrice,    // L: ML단가
        amount,       // M: 금액
        result,       // N: 처리결과
        itemUrl       // O: URL
      ]);
      hasItem = true;
    }

    if (hasItem) {
      newDoneIds.push(orderId);
      newCount++;
    }
  }

  // 재고 반영
  if (deductCount > 0) {
    invSheet.getRange(2,1,invData.length,7).setValues(invData);
    try { invCacheClear(); } catch(e) {}
    try { 통합재고동기화(); } catch(e) {}
  }

  // ML판매내역 로그 기록
  if (logRows.length > 0) {
    var lastLogRow = salesSheet.getLastRow() + 1;
    if (lastLogRow < 2) lastLogRow = 2;

    var logRange = salesSheet.getRange(lastLogRow,1,logRows.length,15);
    logRange.setValues(logRows);

    // A열 체크박스
    salesSheet.getRange(lastLogRow,1,logRows.length,1)
      .insertCheckboxes();

    // B열 날짜 포맷
    salesSheet.getRange(lastLogRow,2,logRows.length,1)
      .setNumberFormat("yyyy-MM-dd HH:mm");

    // H열 ML/FLEX 색상
    for (var lr = 0; lr < logRows.length; lr++) {
      var typeCell = salesSheet.getRange(lastLogRow+lr, 8);
      if (logRows[lr][7] === "FLEX") {
        typeCell.setBackground("#D5E8D4").setFontColor("#27AE60").setFontWeight("bold");
      } else {
        typeCell.setBackground("#DAE8FC").setFontColor("#2980B9").setFontWeight("bold");
      }
    }

    // L열 단가, M열 금액 포맷
    salesSheet.getRange(lastLogRow,12,logRows.length,2).setNumberFormat("#,##0");

    // O열 URL 색상
    salesSheet.getRange(lastLogRow,15,logRows.length,1)
      .setFontColor("#2980B9");

    // C열 주문ID 텍스트 형식
    salesSheet.getRange(lastLogRow,3,logRows.length,1).setNumberFormat("@");
  }

  // 완료주문번호 저장
  // 완료주문번호 저장 (날짜 + ML/FLEX 추가)
  if (newDoneIds.length > 0) {
    var lastDoneRow = doneSheet.getLastRow() + 1;

    // 헤더 없으면 추가
    if (doneSheet.getLastRow() < 1) {
      doneSheet.getRange(1,1,1,3).setValues([["주문번호","날짜","ML/FLEX"]])
        .setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold");
      lastDoneRow = 2;
    }

    // 주문번호별 날짜/타입 맵
    var orderInfoMap = {};
    for (var o = 0; o < orders.length; o++) {
      var rawOrder = JSON.stringify(orders[o]);
      var idMatch  = rawOrder.match(/"id"\s*:\s*(\d{15,})/);
      var oid      = idMatch ? idMatch[1] : String(orders[o].id);
      var oDate    = orders[o].date_created
        ? Utilities.formatDate(new Date(orders[o].date_created),'America/Santiago','yyyy-MM-dd')
        : Utilities.formatDate(new Date(),'America/Santiago','yyyy-MM-dd');

      // ML/FLEX 타입 찾기
      var oType = "ML";
      var shippingId2 = orders[o].shipping ? String(orders[o].shipping.id||"") : "";
      if (shippingId2) {
        try {
          var sr = UrlFetchApp.fetch(
            "https://api.mercadolibre.com/shipments/" + shippingId2,
            {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
          );
          if (sr.getResponseCode() === 200) {
            var sd = JSON.parse(sr.getContentText());
            var lt = String(sd.logistic_type||"").toLowerCase();
            if (lt === "self_service" || lt.indexOf("fulfillment") !== -1 || lt.indexOf("flex") !== -1) {
              oType = "FLEX";
            }
          }
        } catch(e) {}
      }
      orderInfoMap[oid] = {date: oDate, type: oType};
    }

    var doneRows = newDoneIds.map(function(id) {
      var info = orderInfoMap[id] || {date: Utilities.formatDate(new Date(),'America/Santiago','yyyy-MM-dd'), type: "ML"};
      return [id, info.date, info.type];
    });

    doneSheet.getRange(lastDoneRow,1,doneRows.length,3)
      .setValues(doneRows).setNumberFormat("@");
    doneSheet.getRange(lastDoneRow,2,doneRows.length,1)
      .setNumberFormat("yyyy-MM-dd");
  }

  // 결과 메시지
  var msg = "✅ ML 판매 처리완료!\n\n" +
    "신규 주문: "     + newCount    + "건\n" +
    "재고 차감: "     + deductCount + "개 SKU\n" +
    "스킵(기처리): "  + skipCount   + "건\n\n" +
    "ML판매내역 확인 후\n체크박스 선택 → Ventas 전송 클릭";

  if (notFound.length > 0) {
    msg += "\n\n❌ SKU 없음:\n" + notFound.join("\n");
  }
  if (notInBundle.length > 0) {
    msg += "\n\n⚠️ 묶음설정 미등록:\n" + notInBundle.slice(0,5).join("\n");
    if (notInBundle.length > 5) msg += "\n외 " + (notInBundle.length-5) + "개";
  }

  ui.alert(msg);
}

// ── Ventas 전송 (체크된 행만) ────────────────────────────
function mlSalesToVentas() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var mlSheet = ss.getSheetByName("ML판매내역");
  if (!mlSheet || mlSheet.getLastRow() < 2) {
    ui.alert("⚠️ ML판매내역 시트에 데이터가 없어요!"); return;
  }

  var lastRow = mlSheet.getLastRow();
  var data    = mlSheet.getRange(2,1,lastRow-1,15).getValues();

  // 헤더에서 열 위치 파악
  var headers = mlSheet.getRange(1,1,1,15).getValues()[0];
  var colMap  = {};
  for (var h = 0; h < headers.length; h++) {
    colMap[String(headers[h]).trim()] = h;
  }

  var COL_CHK    = colMap["선택"]        !== undefined ? colMap["선택"]        : 0;
  var COL_SKU    = colMap["SKU"]         !== undefined ? colMap["SKU"]         : 5;
  var COL_NOMBRE = colMap["상품명"]      !== undefined ? colMap["상품명"]      : 6;
  var COL_FLEX   = colMap["ML/FLEX"]     !== undefined ? colMap["ML/FLEX"]     : 7;
  var COL_QTY    = colMap["실제차감수량"] !== undefined ? colMap["실제차감수량"] : 10;
  var COL_PRICE  = colMap["ML단가"]      !== undefined ? colMap["ML단가"]      : 11;
  var COL_AMT    = colMap["금액"]        !== undefined ? colMap["금액"]        : 12;

  // 체크된 행만 필터
  var pendingRows = [];
  var pendingIdxs = [];

  for (var i = 0; i < data.length; i++) {
    var checked = data[i][COL_CHK];
    if (!(checked === true || checked === "TRUE")) continue;

    var sku     = String(data[i][COL_SKU]   ||"").trim();
    var nombre  = String(data[i][COL_NOMBRE]||"").trim();
    var mlType  = String(data[i][COL_FLEX]  ||"ML").trim();
    var realQty = Number(data[i][COL_QTY]   ||0);
  var mlPrice = Number(data[i][COL_PRICE] ||0);
    console.log("SKU: " + String(data[i][COL_SKU]||"") + " ML단가: " + mlPrice);
    var amount  = Number(data[i][COL_AMT]   ||0);

    if (!sku || realQty <= 0) continue;

    pendingRows.push({
      idx:     i,
      sku:     sku,
      nombre:  nombre,
      mlType:  mlType,
      realQty: realQty,
      mlPrice: mlPrice,
      amount:  amount
    });
    pendingIdxs.push(i);
  }

  if (pendingRows.length === 0) {
    ui.alert("⚠️ 체크된 항목이 없어요!\nA열 체크박스를 선택해주세요.");
    return;
  }

  // INVENTARIO에서 바코드 조회
  var invSheet   = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invData    = invSheet.getDataRange().getValues();
  var barcodeMap = {};
  var invPriceMap= {};

  for (var j = 1; j < invData.length; j++) {
    var s = String(invData[j][1]||"").trim();
    if (s) {
      barcodeMap[s]   = String(invData[j][0]||"").trim();
      invPriceMap[s]  = Number(invData[j][3])||0;
    }
  }

  // Ventas 시트에 기록
  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (!ventasSheet) {
    ui.alert("⚠️ Ventas 시트를 찾을 수 없어요!"); return;
  }

  var today   = new Date();
  var rows    = [];
  var notFound= [];

  for (var k = 0; k < pendingRows.length; k++) {
    var row     = pendingRows[k];
    var barcode = barcodeMap[row.sku] || "";
    // ML단가 (L열) 우선 사용
    var price   = row.mlPrice > 0 ? row.mlPrice : (invPriceMap[row.sku]||0);
    var importe = price * row.realQty;

    if (!barcode) notFound.push(row.sku);

    rows.push([
      today,         // A: Fecha
      "",            // B: Cliente (수동입력)
      "",            // C: Tipo de pago (수동입력)
      barcode,       // D: Codigo de barras
      row.sku,       // E: Codigo SKU
      row.nombre,    // F: Nombre del producto
      row.realQty,   // G: Cantidad
      price,         // H: Precio
      importe,       // I: Importe
      ""             // J: Total general
    ]);
  }

 if (rows.length === 0) {
    ui.alert("⚠️ 전송할 데이터가 없어요!"); return;
  }

  var ventasLastRow = ventasSheet.getLastRow() + 1;
  if (ventasLastRow < 2) ventasLastRow = 2;

  // 디버그 확인
  ui.alert(
    "Ventas 전송 직전 확인\n\n" +
    "전송할 행 수: " + rows.length + "\n" +
    "Ventas 마지막행: " + ventasSheet.getLastRow() + "\n" +
    "입력 시작행: " + ventasLastRow + "\n\n" +
    "첫번째 행 데이터:\n" +
    "날짜: " + rows[0][0] + "\n" +
    "거래처: " + rows[0][1] + "\n" +
    "SKU: " + rows[0][4] + "\n" +
    "수량: " + rows[0][6] + "\n" +
    "단가: " + rows[0][7] + "\n" +
    "금액: " + rows[0][8]
  );

 // 데이터 유효성 검사 무시하고 입력
  ventasSheet.getRange(ventasLastRow,1,rows.length,10).setValues(rows);
  ventasSheet.getRange(ventasLastRow,1,rows.length,1)
    .setNumberFormat("yyyy-MM-dd HH:mm");
  ventasSheet.getRange(ventasLastRow,8,rows.length,2)
    .setNumberFormat("#,##0");

  try { actualizarTotal(ventasSheet); } catch(e) {}

  // 체크된 행 삭제 (아래→위 순서)
  var deleteRows = pendingIdxs.map(function(idx){ return idx+2; });
  deleteRows.sort(function(a,b){ return b-a; });
  for (var dr = 0; dr < deleteRows.length; dr++) {
    mlSheet.deleteRow(deleteRows[dr]);
  }

  // Ventas로 이동
  ss.setActiveSheet(ventasSheet);
  ventasSheet.setActiveRange(ventasSheet.getRange(ventasLastRow,2));

  var msg = "✅ Ventas 전송완료!\n\n" +
    "전송: "  + rows.length      + "개 항목\n" +
    "삭제: "  + deleteRows.length + "행 (ML판매내역)\n\n" +
    "판매완료 버튼 클릭하세요!";
  if (notFound.length > 0) {
    msg += "\n\n⚠️ 바코드 없는 SKU:\n" + notFound.join("\n");
  }
  ui.alert(msg);
}

// ── ML 판매내역 수동 입력 (API 없이) ────────────────────
// API 오류시 수동으로 ML 판매수량을 입력해서 재고 차감
function mlSalesManualDeduct() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var bundleSheet = ss.getSheetByName("묶음설정");
  if (!bundleSheet || bundleSheet.getLastRow() < 2) {
    ui.alert("⚠️ 묶음설정 시트를 먼저 만들어주세요!");
    return;
  }

  // ML수동입력 시트 확인
  var manualSheet = ss.getSheetByName("ML수동입력");
  if (!manualSheet) {
    manualSheet = ss.insertSheet("ML수동입력");
    var hdrs2 = ["SKU", "상품명", "ML판매수량", "묶음수량", "실제차감수량", "처리결과"];
    manualSheet.getRange(1, 1, 1, hdrs2.length)
      .setValues([hdrs2])
      .setBackground("#3a2a1a").setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center");
    manualSheet.setFrozenRows(1);

    // 묶음설정에서 SKU 목록 자동 채우기
    var bundleData2 = bundleSheet.getRange(2, 1, bundleSheet.getLastRow() - 1, 4).getValues();
    var rows2 = [];
    for (var b = 0; b < bundleData2.length; b++) {
      var sku2 = String(bundleData2[b][0] || "").trim();
      var nm2  = String(bundleData2[b][1] || "").trim();
      var bun2 = Number(bundleData2[b][3]) || 1;
      if (sku2) rows2.push([sku2, nm2, 0, bun2, "", ""]);
    }
    if (rows2.length > 0) {
      manualSheet.getRange(2, 1, rows2.length, 6).setValues(rows2);
    }
    ui.alert(
      "✅ ML수동입력 시트 생성!\n\n" +
      "C열(ML판매수량)에 오늘 ML에서 판매된 수량 입력 후\n" +
      "메뉴 → ML 판매자동화 → 수동입력 재고 차감 실행"
    );
    ss.setActiveSheet(manualSheet);
    return;
  }

  // 수동입력 데이터 처리
  var lastRow = manualSheet.getLastRow();
  if (lastRow < 2) { ui.alert("데이터가 없어요!"); return; }

  var manualData = manualSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invData2 = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 7).getValues();
  var invMap2 = {};
  for (var im2 = 0; im2 < invData2.length; im2++) {
    var s2 = String(invData2[im2][1] || "").trim();
    if (s2) invMap2[s2] = im2;
  }

  var count2 = 0;
  var salesLogRows = [];

  for (var r = 0; r < manualData.length; r++) {
    var sku3    = String(manualData[r][0] || "").trim();
    var mlQty2  = Number(manualData[r][2]) || 0;
    var bundle2 = Number(manualData[r][3]) || 1;
    var realQty2 = mlQty2 * bundle2;

    if (!sku3 || mlQty2 <= 0) {
      manualSheet.getRange(r + 2, 5).setValue("");
      manualSheet.getRange(r + 2, 6).setValue("");
      continue;
    }

    if (invMap2.hasOwnProperty(sku3)) {
      var idx2 = invMap2[sku3];
      var cur2 = Number(invData2[idx2][6]) || 0;
      var new2 = Math.max(0, cur2 - realQty2);
      invData2[idx2][6] = new2;

      manualSheet.getRange(r + 2, 5).setValue(realQty2);
      manualSheet.getRange(r + 2, 6)
        .setValue("✅ " + cur2 + " → " + new2)
        .setFontColor("#1a3a2a").setFontWeight("bold");

      // ML판매내역에 기록
      salesLogRows.push([
        new Date(), "수동입력", "-", "-",
        sku3, manualData[r][1], mlQty2, bundle2, realQty2,
        "✅ " + cur2 + " → " + new2
      ]);
      count2++;
    } else {
      manualSheet.getRange(r + 2, 6)
        .setValue("❌ SKU 없음")
        .setFontColor("#e74c3c").setFontWeight("bold");
    }
  }

  // 재고 반영
  invSheet.getRange(2, 1, invData2.length, 7).setValues(invData2);
  invCacheClear();
  통합재고동기화();

  // ML판매내역 시트에 로그
  if (salesLogRows.length > 0) {
    var salesSheet2 = ss.getSheetByName("ML판매내역");
    if (!salesSheet2) {
      salesSheet2 = ss.insertSheet("ML판매내역");
      var hdrs3 = ["처리일시","주문ID","계정","ML_ID","SKU","상품명","ML판매수량","묶음수량","실제차감수량","처리결과"];
      salesSheet2.getRange(1,1,1,hdrs3.length).setValues([hdrs3])
        .setBackground("#1a3a2a").setFontColor("#FFFFFF").setFontWeight("bold");
      salesSheet2.setFrozenRows(1);
    }
    var lastSalesRow = salesSheet2.getLastRow() + 1;
    salesSheet2.getRange(lastSalesRow, 1, salesLogRows.length, 10).setValues(salesLogRows);
    salesSheet2.getRange(lastSalesRow, 1, salesLogRows.length, 1).setNumberFormat("yyyy-MM-dd HH:mm");
  }

  // 수동입력 수량 초기화 (다음날 재사용)
  for (var rr = 0; rr < manualData.length; rr++) {
    manualSheet.getRange(rr + 2, 3).setValue(0);
  }

  ui.alert("✅ ML 수동 재고 차감 완료!\n" + count2 + "개 SKU 처리됨");
}

function 묶음설정시트열기() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = 묶음설정시트확인();
  ss.setActiveSheet(sheet);
}

function ml수동입력시트열기() {
  mlSalesManualDeduct();
}

function testMLOrdersCheck() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getMLTokenSheet();
  
  var token    = String(sheet.getRange(3, 2).getValue() || "").trim();
  var sellerId = String(sheet.getRange(5, 2).getValue() || "").trim();
  
  if (!token) { ui.alert("❌ ACCESS_TOKEN이 없어요!"); return; }
  if (!sellerId) { ui.alert("❌ SELLER_ID가 없어요!"); return; }
  
  // 1. 토큰 유효성 확인
  var meResp = UrlFetchApp.fetch("https://api.mercadolibre.com/users/me", {
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
  
  var meCode = meResp.getResponseCode();
  var meData = JSON.parse(meResp.getContentText());
  
  if (meCode !== 200) {
    ui.alert("❌ 토큰 오류(" + meCode + ")\n\n" + meResp.getContentText());
    return;
  }
  
  // 2. 최근 주문 조회
  var ordResp = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=" + sellerId +
    "&sort=date_desc&limit=5",
    {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    }
  );
  
  var ordCode = ordResp.getResponseCode();
  var ordText = ordResp.getContentText();
  
  var msg = "✅ 토큰 유효\n";
  msg += "계정: " + (meData.nickname || meData.id) + "\n";
  msg += "SELLER_ID: " + sellerId + "\n\n";
  msg += "주문조회 응답코드: " + ordCode + "\n";
  
  if (ordCode === 200) {
    var ordData = JSON.parse(ordText);
    var total = ordData.paging ? ordData.paging.total : 0;
    var results = ordData.results || [];
    msg += "전체 주문수: " + total + "건\n";
    msg += "가져온 주문: " + results.length + "건\n";
    if (results.length > 0) {
      msg += "\n최근 주문 ID: " + results[0].id;
      msg += "\n주문 상태: " + results[0].status;
      msg += "\n주문 날짜: " + results[0].date_created;
    }
  } else {
    msg += "주문조회 오류:\n" + ordText.substring(0, 200);
  }
  
  ui.alert(msg);
}

function testMLOrdersCheck() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getMLTokenSheet();
  
  var accessToken = String(sheet.getRange(4, 2).getValue() || "").trim();
  var sellerId    = String(sheet.getRange(6, 2).getValue() || "").trim();
  
  // 토큰 유효성 확인
  var meResp = UrlFetchApp.fetch("https://api.mercadolibre.com/users/me", {
    headers: { "Authorization": "Bearer " + accessToken },
    muteHttpExceptions: true
  });
  
  if (meResp.getResponseCode() !== 200) {
    ui.alert("❌ 토큰 오류\n" + meResp.getContentText());
    return;
  }
  
  var meData = JSON.parse(meResp.getContentText());
  
  // 주문 조회
  var ordResp = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=" + sellerId +
    "&sort=date_desc&limit=5",
    {
      headers: { "Authorization": "Bearer " + accessToken },
      muteHttpExceptions: true
    }
  );
  
  var ordCode = ordResp.getResponseCode();
  var msg = "✅ 토큰 유효\n";
  msg += "계정: " + (meData.nickname || meData.id) + "\n";
  msg += "SELLER_ID: " + sellerId + "\n\n";
  msg += "주문조회 응답코드: " + ordCode + "\n";
  
  if (ordCode === 200) {
    var ordData = JSON.parse(ordResp.getContentText());
    var total   = ordData.paging ? ordData.paging.total : 0;
    var results = ordData.results || [];
    msg += "전체 주문수: " + total + "건\n";
    msg += "최근 주문: "   + results.length + "건\n";
    if (results.length > 0) {
      var last = results[0];
      msg += "\n최근 주문 ID: "  + last.id;
      msg += "\n주문 상태: "     + last.status;
      msg += "\n주문 날짜: "     + last.date_created;
      msg += "\n상품수: "        + (last.order_items ? last.order_items.length : 0) + "개";
    }
  } else {
    msg += "오류: " + ordResp.getContentText().substring(0, 300);
  }
  
  ui.alert(msg);
}


function addButtonToMLSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ML판매내역");
  if (!sheet) { 
    SpreadsheetApp.getUi().alert("ML판매내역 시트가 없어요!"); 
    return; 
  }
  
  // L1셀에 버튼 안내 텍스트
  sheet.getRange("L1")
    .setValue("👇 아래 버튼 클릭")
    .setFontWeight("bold")
    .setFontColor("#1a3a2a");
  
  sheet.getRange("L2")
    .setValue("▶ Ventas로 전송")
    .setBackground("#1a3a2a")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  SpreadsheetApp.getUi().alert(
    "✅ 안내셀 추가완료!\n\n" +
    "ML판매내역 시트 L2셀에\n" +
    "버튼 모양을 그림으로 추가하려면:\n\n" +
    "삽입 → 그림 → 버튼 그리기 후\n" +
    "스크립트: mlSalesToVentas 연결"
  );
}

function addBarcodeCheckToMLSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var mlSheet  = ss.getSheetByName("ML판매내역");
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);

  if (!mlSheet || !invSheet) {
    ui.alert("시트를 찾을 수 없어요!");
    return;
  }

  // INVENTARIO에서 SKU → 바코드 맵
  var invData    = invSheet.getDataRange().getValues();
  var barcodeMap = {};
  var skuMap     = {}; // SKU → 상품명도 같이
  for (var j = 1; j < invData.length; j++) {
    var s = String(invData[j][1] || "").trim();
    if (s) {
      barcodeMap[s] = String(invData[j][0] || "").trim();
      skuMap[s]     = String(invData[j][2] || "").trim();
    }
  }

  // 헤더 추가 (K, L, M열)
  mlSheet.getRange("K1").setValue("바코드(자동조회)")
    .setBackground("#1a3a2a").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");
  mlSheet.getRange("L1").setValue("SKU매칭여부")
    .setBackground("#1a3a2a").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");
  mlSheet.getRange("M1").setValue("INVENTARIO상품명")
    .setBackground("#1a3a2a").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");

  mlSheet.setColumnWidth(11, 160); // K열
  mlSheet.setColumnWidth(12, 110); // L열
  mlSheet.setColumnWidth(13, 200); // M열

  // 데이터 행 처리
  var lastRow = mlSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("ML판매내역 데이터가 없어요!");
    return;
  }

  var mlData = mlSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var found = 0, notFound = 0;

  for (var i = 0; i < mlData.length; i++) {
    var sku     = String(mlData[i][4] || "").trim(); // E열 SKU
    var rowNum  = i + 2;

    if (!sku) continue;

    if (barcodeMap[sku]) {
      // 매칭 성공
      mlSheet.getRange(rowNum, 11).setValue(barcodeMap[sku]); // K: 바코드
      mlSheet.getRange(rowNum, 12)
        .setValue("✅ 일치")
        .setFontColor("#1a3a2a")
        .setFontWeight("bold")
        .setBackground("#e8f5e9");
      mlSheet.getRange(rowNum, 13).setValue(skuMap[sku]);     // M: 상품명
      found++;
    } else {
      // 매칭 실패
      mlSheet.getRange(rowNum, 11).setValue(""); 
      mlSheet.getRange(rowNum, 12)
        .setValue("❌ 없음")
        .setFontColor("#e74c3c")
        .setFontWeight("bold")
        .setBackground("#fadbd8");
      mlSheet.getRange(rowNum, 13)
        .setValue("← SKU 확인필요")
        .setFontColor("#e74c3c");
      notFound++;
    }
  }

  ss.setActiveSheet(mlSheet);
  ui.alert(
    "✅ 바코드 확인 완료!\n\n" +
    "✅ 매칭됨: " + found + "개\n" +
    "❌ 못찾음: " + notFound + "개\n\n" +
    "❌ 항목은 E열 SKU를 확인해서\n" +
    "묶음설정 시트 A열 SKU와\n" +
    "정확히 맞춰주세요!"
  );
}

// ========================
// 거래처관리 파일 연동
// ========================
 var GEORAE_FILE_ID = '1H9nj7374d6aRQknyy78phwq5RwE_4M2lRSJoH0vB3J4';// ← 거래처관리 파일 ID 입력

function inventarioToGeorae() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var row   = sheet.getActiveRange().getRow();

  if (row < 2) { SpreadsheetApp.getUi().alert("상품 행을 선택하세요!"); return; }

  var sku    = String(sheet.getRange(row, 2).getValue() || "").trim();
  var nameEs = String(sheet.getRange(row, 3).getValue() || "").trim();
  if (!sku) { SpreadsheetApp.getUi().alert("SKU가 없는 행이에요!"); return; }

  // H열 이미지 URL 추출
  var imgFormula = sheet.getRange(row, 8).getFormula();
  var imgUrl     = "";
  if (imgFormula) {
    var m = imgFormula.match(/IMAGE\s*\(\s*["']([^"']+)["']/i);
    if (m && m[1]) imgUrl = m[1];
  }
  if (!imgUrl) {
    var val = String(sheet.getRange(row, 8).getValue() || "").trim();
    if (val.indexOf("http") === 0) imgUrl = val;
  }

  // 거래처관리 파일 _임시데이터에 저장
  try {
    var georaeSS  = SpreadsheetApp.openById(GEORAE_FILE_ID);
    var tempSheet = georaeSS.getSheetByName("_임시데이터");
    if (!tempSheet) {
      SpreadsheetApp.getUi().alert("거래처관리 파일에 _임시데이터 시트가 없어요!");
      return;
    }
    tempSheet.getRange("A1").setValue(sku);
    tempSheet.getRange("B1").setValue(nameEs);
    tempSheet.getRange("C1").setValue(imgUrl);
    SpreadsheetApp.flush();
  } catch(err) {
    SpreadsheetApp.getUi().alert("오류: " + err.message);
    return;
  }

  // 거래처관리 파일에서 직접 팝업 실행
  try {
    var georaeSS2 = SpreadsheetApp.openById(GEORAE_FILE_ID);

    // 거래처관리 파일의 함수를 직접 실행
    // → NAOS 재고에서 팝업 HTML을 직접 생성해서 보여줌
    _openGeoraePopupFromNAOS(sku, nameEs, imgUrl);

  } catch(err) {
    // 직접 실행 안되면 파일 열기로 fallback
    var georaeUrl = "https://docs.google.com/spreadsheets/d/" + GEORAE_FILE_ID + "/edit";
    var html = HtmlService.createHtmlOutput(
      '<script>window.onload=function(){window.open("' + georaeUrl + '","_blank");google.script.host.close();};</script>' +
      '<div style="font-family:Arial;padding:20px;text-align:center;color:#784212;">' +
      '<div style="font-size:13px;font-weight:bold;">거래처관리 파일을 열고 있어요...</div>' +
      '<div style="font-size:11px;color:#888;margin-top:6px;">파일이 열리면 메뉴에서<br>⚡ 제품+거래처 통합 등록을 클릭하세요</div></div>'
    ).setWidth(280).setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(html, "거래처관리 연동");
  }
}

function _openGeoraePopupFromNAOS(sku, nameEs, imgUrl) {
  var displayUrl = "";
  if (imgUrl) {
    if (imgUrl.indexOf('lh3.googleusercontent.com') !== -1) {
      displayUrl = imgUrl;
    } else if (imgUrl.indexOf('/d/') !== -1) {
      var idM = imgUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      displayUrl = idM ? 'https://drive.google.com/thumbnail?id=' + idM[1] + '&sz=w150' : imgUrl;
    } else {
      displayUrl = imgUrl;
    }
  }

  var imgUrlB64     = Utilities.base64Encode(displayUrl || '');
  var imgFormulaB64 = Utilities.base64Encode(imgUrl     || '');

  // 업체 입력박스 HTML 생성 함수
  function supplierBox(num, required) {
    var req = required ? ' *' : '';
    return '<div class="sbox" id="sbox' + num + '">'
      + '<div class="stitle"><span class="snum">' + num + '</span> 업체' + req + '</div>'
      + '<div class="r2">'
      + '<div><label>업체명' + req + '</label><input id="s' + num + '_name" placeholder="예: 광저우철물"></div>'
      + '<div><label>전화번호</label><input id="s' + num + '_phone" placeholder="138-0000-0000"></div>'
      + '</div>'
      + '<div class="r3">'
      + '<div><label>단가 (CNY)</label><input id="s' + num + '_price" placeholder="0.93" type="number" step="0.001"></div>'
      + '<div><label>구매년월</label><input id="s' + num + '_date" type="month"></div>'
      + '<div><label>메모</label><input id="s' + num + '_memo" placeholder="100pcs/box"></div>'
      + '</div>'
      + '<div><label>URL (1688 / 알리바바)</label>'
      + '<input id="s' + num + '_url" placeholder="https://detail.1688.com/..."></div>'
      + '</div>';
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>'
    + '*{box-sizing:border-box;margin:0;padding:0;}'
    + 'body{font-family:Arial,sans-serif;background:#FEF9F4;color:#784212;padding:14px;font-size:12px;}'
    + 'h2{font-size:13px;font-weight:700;margin:0 0 10px;padding-bottom:6px;border-bottom:2px solid #F0B27A;}'
    + '.sec{font-size:11px;font-weight:700;margin:10px 0 6px;padding:4px 8px;'
    + 'background:#FAD7A0;border-radius:4px;color:#784212;}'
    + '.r2{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:6px;}'
    + '.r3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:6px;}'
    + 'label{font-size:10px;color:#A04000;display:block;margin-bottom:2px;font-weight:600;}'
    + 'input,select{width:100%;padding:5px 8px;background:#FFFBF7;border:1.5px solid #F0B27A;'
    + 'border-radius:5px;color:#784212;font-size:11px;outline:none;}'
    + 'input:focus,select:focus{border-color:#C0392B;background:#fff;}'
    + '.sbox{background:#FFFBF7;border:1.5px solid #F0B27A;border-radius:8px;'
    + 'padding:10px;margin-bottom:8px;}'
    + '.stitle{font-size:11px;font-weight:700;color:#C0392B;margin-bottom:8px;'
    + 'padding-bottom:4px;border-bottom:1px dashed #FAD7A0;'
    + 'display:flex;align-items:center;gap:6px;}'
    + '.snum{background:#C0392B;color:#fff;border-radius:50%;'
    + 'width:18px;height:18px;display:flex;align-items:center;'
    + 'justify-content:center;font-size:10px;font-weight:bold;flex-shrink:0;}'
    + '.img-row{display:flex;gap:10px;align-items:center;margin-bottom:10px;}'
    + '.prod-sku{font-size:14px;font-weight:bold;color:#C0392B;}'
    + '.prod-name{font-size:11px;color:#784212;margin-top:3px;}'
    + '.br{display:flex;gap:8px;margin-top:12px;justify-content:flex-end;}'
    + 'button{padding:8px 20px;border-radius:6px;border:none;'
    + 'font-size:12px;font-weight:700;cursor:pointer;}'
    + '.sv{background:#E59866;color:#fff;}'
    + '.cl{background:#FEF5E7;color:#784212;border:1.5px solid #F0B27A;}'
    + '.mg{font-size:11px;margin-top:8px;text-align:center;min-height:16px;color:#E59866;}'
    + '</style></head><body>'

    // 타이틀
    + '<h2>⚡ 제품 + 거래처 통합 등록</h2>'

    // 사진 + SKU
    + '<div class="img-row">'
    + '<div style="flex-shrink:0;">'
    + (displayUrl
        ? '<img src="' + displayUrl + '" style="width:75px;height:75px;object-fit:contain;'
          + 'border-radius:6px;border:1.5px solid #F0B27A;">'
        : '<div style="width:75px;height:75px;background:#f5f5f5;border-radius:6px;'
          + 'display:flex;align-items:center;justify-content:center;'
          + 'color:#bbb;font-size:10px;border:1.5px solid #ddd;">사진없음</div>')
    + '</div>'
    + '<div>'
    + '<div class="prod-sku">' + sku + '</div>'
    + '<div class="prod-name">' + nameEs + '</div>'
    + '</div></div>'

    // 제품 정보
    + '<div class="sec">📋 제품 정보</div>'
    + '<input type="hidden" id="sku" value="' + sku + '">'
    + '<div class="r2">'
    + '<div><label>상품명 (스페인어) *</label>'
    + '<input id="es" value="' + nameEs + '"></div>'
    + '<div><label>상품명 (한국어)</label>'
    + '<input id="ko" placeholder="경첩 150mm"></div>'
    + '</div>'
    + '<div class="r2">'
    + '<div><label>상품명 (중국어) — 발주서 표시</label>'
    + '<input id="zh" placeholder="铰链 150mm"></div>'
    + '<div><label>품목</label>'
    + '<input id="item" placeholder="경첩 / 비사그라"></div>'
    + '</div>'
    + '<div class="r2">'
    + '<div><label>규격</label>'
    + '<input id="sz" placeholder="150mm / 90g"></div>'
    + '<div><label>ML 여부</label>'
    + '<select id="ml">'
    + '<option value="X 미등록">X 미등록</option>'
    + '<option value="O 등록">O 등록</option>'
    + '</select></div>'
    + '</div>'
    + '<div style="margin-bottom:6px;">'
    + '<label>URL</label>'
    + '<input id="url" placeholder="https://..."></div>'

    // 거래처 3개
    + '<div class="sec">💰 거래처 단가 (업체 최대 3개)</div>'
    + supplierBox(1, true)
    + supplierBox(2, false)
    + supplierBox(3, false)

    // 버튼
    + '<div class="br">'
    + '<button class="cl" onclick="google.script.host.close()">취소</button>'
    + '<button class="sv" onclick="go()">💾 저장</button>'
    + '</div>'
    + '<div class="mg" id="mg"></div>'
    + '<input type="hidden" id="iub" value="' + imgUrlB64 + '">'
    + '<input type="hidden" id="ifb" value="' + imgFormulaB64 + '">'

    + '<script>'
    + 'function v(id){return document.getElementById(id);}'
    + 'function b64d(s){try{return decodeURIComponent(escape(atob(s)));}catch(e){return "";}}'
    + 'function go(){'
    + '  var sku=v("sku").value.trim();'
    + '  var es=v("es").value.trim();'
    + '  if(!sku||!es){'
    + '    v("mg").style.color="#C0392B";'
    + '    v("mg").innerText="SKU와 스페인어 상품명은 필수입니다";'
    + '    return;'
    + '  }'
    + '  if(!v("s1_name").value.trim()){'
    + '    v("mg").style.color="#C0392B";'
    + '    v("mg").innerText="업체 1 업체명은 필수입니다";'
    + '    return;'
    + '  }'
    + '  var suppliers=[];'
    + '  for(var i=1;i<=3;i++){'
    + '    var nm=v("s"+i+"_name").value.trim();'
    + '    if(!nm) continue;'
    + '    suppliers.push({'
    + '      sname:  nm,'
    + '      sphone: v("s"+i+"_phone").value.trim(),'
    + '      p1:     v("s"+i+"_price").value.trim(),'
    + '      d1:     v("s"+i+"_date").value.trim(),'
    + '      m1:     v("s"+i+"_memo").value.trim(),'
    + '      u1:     v("s"+i+"_url").value.trim()'
    + '    });'
    + '  }'
    + '  var d={'
    + '    sku:sku, es:es,'
    + '    ko:      v("ko").value.trim(),'
    + '    zh:      v("zh").value.trim(),'
    + '    item:    v("item").value.trim(),'
    + '    sz:      v("sz").value.trim(),'
    + '    ml:      v("ml").value,'
    + '    url:     v("url").value.trim(),'
    + '    price:"", phone:"", note:"",'
    + '    imgUrl:    b64d(v("iub").value),'
    + '    imgFormula:b64d(v("ifb").value),'
    + '    suppliers: suppliers'
    + '  };'
    + '  v("mg").style.color="#E59866";'
    + '  v("mg").innerText="저장 중... 업체 "+suppliers.length+"개";'
    + '  google.script.run'
    + '    .withSuccessHandler(function(msg){'
    + '      v("mg").style.color="#27AE60";'
    + '      v("mg").innerText=msg||"✅ 저장완료!";'
    + '      setTimeout(function(){google.script.host.close();},1000);'
    + '    })'
    + '    .withFailureHandler(function(e){'
    + '      v("mg").style.color="#C0392B";'
    + '      v("mg").innerText="오류: "+e.message;'
    + '    })'
    + '    .saveGeoraeFromNAOS(d);'
    + '}'
    + '<\/script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(580).setHeight(780),
    '⚡ 제품+거래처 통합 등록'
  );
}

function saveGeoraeFromNAOS(d) {
  var georaeSS = SpreadsheetApp.openById(GEORAE_FILE_ID);
  var h = 70;

  var C_LOCAL = {
    ODD: 'FDFEFE', EVEN: 'FDF2E9',
    HEADER: 'FEF5E7', TEXT: '784212',
    SKU: 'C0392B'
  };
  var SHEET_LOCAL = { MASTER: '제품정보', SUPPLIER: '거래처가격' };

  // ── 제품정보 저장 ──────────────────────────
  var mSh = georaeSS.getSheetByName(SHEET_LOCAL.MASTER);
  if (!mSh) throw new Error('제품정보 시트 없음');

  var mr = -1, lastM = 2;
  var mVals = mSh.getRange(3, 2, 500, 1).getValues();
  for (var i = 0; i < mVals.length; i++) {
    var bm = mVals[i][0].toString().trim();
    if (!bm) continue;
    lastM = i + 3;
    if (bm === d.sku) { mr = i + 3; break; }
  }
  if (mr === -1) mr = lastM + 1;
  if (mr < 3) mr = 3;

  mSh.getRange(mr, 1, 1, 13).setValues([[
    '', d.sku, d.es,
    d.ko   || '',
    d.item || '',
    d.sz   || '',
    d.price|| '', '', '',
    d.url  || '', '',
    d.ml   || 'X 미등록',
    d.note || ''
  ]]);

  if (d.imgFormula && d.imgFormula.indexOf('http') === 0) {
    mSh.getRange(mr, 1).setFormula('=IMAGE("' + d.imgFormula + '")');
  }
  mSh.getRange(mr, 9).setFormula(
    '=IFERROR(ROUND(G' + mr + '*H' + mr + ',2),"")');

  var mbg = (mr % 2 === 1) ? C_LOCAL.ODD : C_LOCAL.EVEN;
  mSh.getRange(mr, 1, 1, 13)
    .setBackground(mbg).setFontColor('#1A252F')
    .setFontSize(11).setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  mSh.getRange(mr, 3, 1, 4).setHorizontalAlignment('left');
  mSh.getRange(mr, 2).setFontColor('#' + C_LOCAL.SKU).setFontWeight('bold');
  mSh.setRowHeightsForced(mr, 1, h);

  // ── 거래처가격 저장 (업체 최대 3개) ──────────
  var sSh = georaeSS.getSheetByName(SHEET_LOCAL.SUPPLIER);
  if (!sSh) throw new Error('거래처가격 시트 없음');

  var suppliers = d.suppliers || [];
  if (suppliers.length === 0) return '✅ 제품정보만 저장됨 (업체 없음)';

  var savedCount = 0;

  for (var s = 0; s < suppliers.length; s++) {
    var sup = suppliers[s];
    if (!sup.sname) continue;

    // 기존 행 찾기 (SKU + 업체명 일치 → 업데이트)
    var sr = -1, lastS = 2;
    var sVals = sSh.getRange(3, 2, 500, 4).getValues();
    for (var j = 0; j < sVals.length; j++) {
      var bs = sVals[j][0].toString().trim();
      if (!bs) continue;
      lastS = j + 3;
      if (bs === d.sku &&
          sVals[j][3].toString().trim() === sup.sname) {
        sr = j + 3; break;
      }
    }
    if (sr === -1) sr = lastS + 1;
    if (sr < 3) sr = 3;

    sSh.getRange(sr, 1, 1, 11).setValues([[
      '',
      d.sku,
      d.ko || d.es,
      d.zh || '',
      sup.sname,
      sup.sphone || '',
      sup.p1     || '',
      sup.d1     || '',
      sup.u1     || '',
      sup.m1     || '',
      false
    ]]);

    if (d.imgFormula && d.imgFormula.indexOf('http') === 0) {
      sSh.getRange(sr, 1).setFormula('=IMAGE("' + d.imgFormula + '")');
    }

    var sbg = (sr % 2 === 1) ? C_LOCAL.ODD : C_LOCAL.EVEN;
    sSh.getRange(sr, 1, 1, 11)
      .setBackground(sbg).setFontColor('#1A252F')
      .setFontSize(11).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sSh.getRange(sr, 3, 1, 2).setHorizontalAlignment('left');
    sSh.getRange(sr, 5)
      .setBackground('#' + C_LOCAL.HEADER)
      .setFontColor('#' + C_LOCAL.TEXT)
      .setFontWeight('bold');
    sSh.getRange(sr, 2)
      .setFontColor('#' + C_LOCAL.SKU)
      .setFontWeight('bold');

    var dvC = SpreadsheetApp.newDataValidation()
      .requireCheckbox().build();
    sSh.getRange(sr, 11).setDataValidation(dvC).setValue(false);
    sSh.setRowHeightsForced(sr, 1, h);

    savedCount++;
  }

  return '✅ 저장완료!  제품정보 1건 + 거래처 ' + savedCount + '개 업체';
}

function testImgAccess() {
  var ui = SpreadsheetApp.getUi();

  // 아까 확인된 파일 ID
  var fileId = '1Fy6c0rC0ZA_IphDRicV8x7S0dzhao_pS';

  try {
    // 방법 1: DriveApp으로 직접 접근
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    ui.alert(
      '✅ DriveApp 성공!\n\n' +
      '파일명: ' + file.getName() + '\n' +
      '크기: ' + blob.getBytes().length + ' bytes\n' +
      '타입: ' + blob.getContentType()
    );
  } catch(e1) {
    // 방법 2: URL fetch 시도
    try {
      var url = 'https://lh3.googleusercontent.com/d/' + fileId;
      var resp = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
      ui.alert(
        '❌ DriveApp 실패\n' +
        '✅ URL fetch 응답코드: ' + resp.getResponseCode() + '\n\n' +
        '오류: ' + e1.message
      );
    } catch(e2) {
      ui.alert(
        '❌ 둘 다 실패\n\n' +
        'DriveApp 오류: ' + e1.message + '\n' +
        'URL fetch 오류: ' + e2.message
      );
    }
  }
}

function setNewMLToken() {
  var ui          = SpreadsheetApp.getUi();
  var code        = "TG-6a0780bbd78a7c000101cda7-323449286";
  var sheet       = getMLTokenSheet();
  var clientId    = String(sheet.getRange(2, 2).getValue() || "").trim();
  var clientSecret= String(sheet.getRange(3, 2).getValue() || "").trim();

  if (!clientId || !clientSecret) {
    ui.alert("❌ CLIENT_ID 또는 CLIENT_SECRET이 없어요!\nML토큰설정 시트를 확인해주세요.");
    return;
  }

  try {
    var response = UrlFetchApp.fetch("https://api.mercadolibre.com/oauth/token", {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: "grant_type=authorization_code" +
               "&client_id="     + clientId +
               "&client_secret=" + clientSecret +
               "&code="          + code +
               "&redirect_uri=https://www.google.com",
      muteHttpExceptions: true
    });

    var code2 = response.getResponseCode();
    var body  = response.getContentText();

    if (code2 !== 200) {
      ui.alert("❌ 토큰 발급 실패 (" + code2 + ")\n\n" + body);
      return;
    }

    var result     = JSON.parse(body);
    var expiryTime = new Date().getTime() + (result.expires_in * 1000);

    sheet.getRange(4, 2).setValue(result.access_token);
    sheet.getRange(5, 2).setValue(result.refresh_token);
    sheet.getRange(7, 2).setValue(expiryTime);

    ui.alert(
      "✅ 토큰 발급 완료!\n\n" +
      "ACCESS_TOKEN: " + result.access_token.substring(0,20) + "...\n" +
      "만료시간: " + result.expires_in + "초 후"
    );

  } catch(e) {
    ui.alert("❌ 오류: " + e.message);
  }
}
function testMLApi403() {
  var ui      = SpreadsheetApp.getUi();
  var sheet   = getMLTokenSheet();
  var token   = String(sheet.getRange(4, 2).getValue() || "").trim();
  var sellerId= String(sheet.getRange(6, 2).getValue() || "").trim();

  // 1. 토큰으로 내 계정 확인
  var meResp = UrlFetchApp.fetch("https://api.mercadolibre.com/users/me", {
    headers: {"Authorization": "Bearer " + token},
    muteHttpExceptions: true
  });
  var meData = JSON.parse(meResp.getContentText());

  var msg = "계정 확인:\n";
  msg += "응답코드: " + meResp.getResponseCode() + "\n";
  msg += "실제 ID: " + (meData.id || "없음") + "\n";
  msg += "닉네임: " + (meData.nickname || "없음") + "\n";
  msg += "시트 SELLER_ID: " + sellerId + "\n\n";

  // ID 일치 여부 확인
  if (meData.id && String(meData.id) !== sellerId) {
    msg += "⚠️ ID 불일치!\n";
    msg += "시트: " + sellerId + "\n";
    msg += "실제: " + meData.id + "\n\n";
    msg += "→ SELLER_ID를 " + meData.id + " 로 수정해야 해요!";
  } else {
    msg += "✅ ID 일치\n\n";

    // 2. 주문 조회 테스트
    var ordResp = UrlFetchApp.fetch(
      "https://api.mercadolibre.com/orders/search?seller=" + sellerId +
      "&sort=date_desc&limit=3",
      {
        headers: {"Authorization": "Bearer " + token},
        muteHttpExceptions: true
      }
    );
    msg += "주문조회 응답코드: " + ordResp.getResponseCode() + "\n";
    if (ordResp.getResponseCode() === 200) {
      var ordData = JSON.parse(ordResp.getContentText());
      msg += "전체 주문수: " + (ordData.paging ? ordData.paging.total : 0) + "건";
    } else {
      msg += "오류: " + ordResp.getContentText().substring(0, 100);
    }
  }

  ui.alert(msg);
}

function restoreDoneOrderIds() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = SpreadsheetApp.getActiveSpreadsheet();

  var token = getMLToken("LUCAS");
  if (!token) { ui.alert("❌ 토큰 없음"); return; }

  // 최근 30일 주문 가져오기
  var dateFrom = new Date();
  dateFrom.setDate(dateFrom.getDate() - 30);
  var dateFromStr = dateFrom.toISOString();

  var url = "https://api.mercadolibre.com/orders/search" +
            "?seller=323449286" +
            "&order.status=paid" +
            "&order.date_created.from=" + dateFromStr +
            "&sort=date_desc&limit=50";

  var response = UrlFetchApp.fetch(url, {
    headers: {"Authorization": "Bearer " + token},
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    ui.alert("❌ API 오류(" + response.getResponseCode() + ")");
    return;
  }

  var rawText = response.getContentText();
  var orders  = JSON.parse(rawText).results || [];

  if (orders.length === 0) {
    ui.alert("주문 없음");
    return;
  }

  // 오늘 날짜
  var today    = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd");

  // 완료주문번호 시트 준비
  var doneSheet = ss.getSheetByName("✅완료주문번호");
  if (!doneSheet) {
    doneSheet = ss.insertSheet("✅완료주문번호");
    doneSheet.getRange(1,1).setValue("주문번호")
      .setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold");
  } else {
    // 기존 데이터 모두 삭제 후 새로 채우기
    if (doneSheet.getLastRow() > 1) {
      doneSheet.getRange(2,1,doneSheet.getLastRow()-1,1).clearContent();
    }
  }

  var todayOrders  = []; // 오늘 주문 (재처리 필요)
  var prevOrders   = []; // 이전 주문 (완료처리)
  var prevRows     = [];

  for (var o = 0; o < orders.length; o++) {
    var rawOrder = JSON.stringify(orders[o]);
    var idMatch  = rawOrder.match(/"id"\s*:\s*(\d{15,})/);
    var orderId  = idMatch ? idMatch[1] : String(orders[o].id);
    var dateStr  = (orders[o].date_created || "").substring(0, 10);

    if (dateStr === today) {
      todayOrders.push(orderId);
    } else {
      prevOrders.push(orderId);
      prevRows.push([orderId]);
    }
  }

  // 오늘 제외한 이전 주문번호만 완료처리
  if (prevRows.length > 0) {
    doneSheet.getRange(2,1,prevRows.length,1).setValues(prevRows)
      .setNumberFormat("@");
  }

  ui.alert(
    "✅ 완료주문번호 복구!\n\n" +
    "이전 주문 (완료처리): " + prevOrders.length + "건\n" +
    "오늘 주문 (재처리 대상): " + todayOrders.length + "건\n\n" +
    "이제 ML 판매 가져오기를 실행하면\n" +
    "오늘 주문만 가져와요!"
  );
}

function testMLAlternativeEndpoints() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = getMLTokenSheet();
  var token = String(sheet.getRange(4, 2).getValue() || "").trim();
  var msg   = "대안 엔드포인트 테스트\n\n";

  // 방법 1: feedback 엔드포인트
  var r1 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=323449286&feedback.purchase_rating=neutral&limit=1",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법1 (feedback): " + r1.getResponseCode() + "\n";

  // 방법 2: date_last_updated 기준
  var r2 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=323449286&order.date_last_updated.from=2026-05-01T00:00:00.000-04:00&limit=1",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법2 (date_last_updated): " + r2.getResponseCode() + "\n";

  // 방법 3: shipments 엔드포인트
  var r3 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/shipments/search?seller_id=323449286&limit=1",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법3 (shipments): " + r3.getResponseCode() + "\n";

  // 방법 4: billing
  var r4 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/billing/integration/periods?user_id=323449286&limit=1",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법4 (billing): " + r4.getResponseCode() + "\n";

  // 방법 5: pack orders
  var r5 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/packs/search?seller_id=323449286&limit=1",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법5 (packs): " + r5.getResponseCode() + "\n";

  ui.alert(msg);
}

function setNewMLTokenNAOSApp() {
  var ui          = SpreadsheetApp.getUi();
  var code        = "TG-6a07837e0a199800019d0d4c-323449286";
  var clientId    = "3465154046014198";   // NAOS JK Trading 앱
  var clientSecret= "XsCPv5yLeS4XfyaT8aPl5RWQEeq3oh4K"; 

  // NAOS JK Trading Secret Key 확인 필요
  if (!clientSecret) {
    ui.alert(
      "❌ NAOS JK Trading 앱 Secret Key가 없어요!\n\n" +
      "developers.mercadolibre.cl/devcenter\n" +
      "→ NAOS JK Trading 앱 클릭\n" +
      "→ Secret Key 확인 후 코드에 입력해주세요."
    );
    return;
  }

  try {
    var response = UrlFetchApp.fetch("https://api.mercadolibre.com/oauth/token", {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: "grant_type=authorization_code" +
               "&client_id="     + clientId +
               "&client_secret=" + clientSecret +
               "&code="          + code +
               "&redirect_uri=https://www.google.com",
      muteHttpExceptions: true
    });

    var code2 = response.getResponseCode();
    var body  = response.getContentText();

    if (code2 !== 200) {
      ui.alert("❌ 실패 (" + code2 + ")\n\n" + body);
      return;
    }

    var result = JSON.parse(body);
    ui.alert(
      "✅ NAOS JK Trading 앱 토큰 발급 완료!\n\n" +
      "ACCESS_TOKEN: " + result.access_token.substring(0,20) + "...\n\n" +
      "이제 주문 조회 테스트해볼게요!"
    );

    // 토큰으로 주문 조회 즉시 테스트
    var testResp = UrlFetchApp.fetch(
      "https://api.mercadolibre.com/orders/search?seller=323449286&order.status=paid&limit=3",
      {
        headers: {"Authorization": "Bearer " + result.access_token},
        muteHttpExceptions: true
      }
    );

    var testCode = testResp.getResponseCode();
    if (testCode === 200) {
      var testData = JSON.parse(testResp.getContentText());
      ui.alert(
        "🎉 주문 조회 성공!\n\n" +
        "전체 주문수: " + (testData.paging ? testData.paging.total : 0) + "건\n\n" +
        "NAOS JK Trading 앱으로 교체할게요!"
      );

      // ML토큰설정 시트에 저장
      var sheet      = getMLTokenSheet();
      var expiryTime = new Date().getTime() + (result.expires_in * 1000);
      sheet.getRange(2, 2).setValue(clientId);
      sheet.getRange(4, 2).setValue(result.access_token);
      sheet.getRange(5, 2).setValue(result.refresh_token);
      sheet.getRange(7, 2).setValue(expiryTime);

    } else {
      ui.alert(
        "❌ 주문 조회 여전히 " + testCode + "\n\n" +
        testResp.getContentText().substring(0, 200)
      );
    }

  } catch(e) {
    ui.alert("❌ 오류: " + e.message);
  }
}

function checkMissingMLIds() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getMLTokenSheet();
  var token = String(sheet.getRange(4, 2).getValue() || "").trim();

  // 묶음설정에 있는 ML_ID 목록
  var bundleSheet = ss.getSheetByName("묶음설정");
  var bundleIds   = {};
  if (bundleSheet && bundleSheet.getLastRow() >= 2) {
    var bData = bundleSheet.getRange(2, 1, bundleSheet.getLastRow()-1, 5).getValues();
    for (var i = 0; i < bData.length; i++) {
      var mlId = String(bData[i][2] || "").trim();
      if (mlId) bundleIds[mlId] = bData[i][0]; // ML_ID → SKU
    }
  }

  // 오늘 주문 가져오기
  var today    = new Date();
  var dateFrom = new Date(today);
  dateFrom.setHours(0, 0, 0, 0);

  var url = "https://api.mercadolibre.com/orders/search" +
            "?seller=323449286" +
            "&order.status=paid" +
            "&order.date_created.from=" + dateFrom.toISOString() +
            "&sort=date_desc&limit=50";

  var response = UrlFetchApp.fetch(url, {
    headers: {"Authorization": "Bearer " + token},
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    ui.alert("❌ API 오류: " + response.getResponseCode());
    return;
  }

  var orders  = JSON.parse(response.getContentText()).results || [];
  var missing = {}; // ML_ID → 상품명

  for (var o = 0; o < orders.length; o++) {
    var items = orders[o].order_items || [];
    for (var oi = 0; oi < items.length; oi++) {
      var item   = items[oi];
      var mlId   = item.item ? String(item.item.id) : "";
      var nombre = item.item ? (item.item.title || "") : "";
      var qty    = Number(item.quantity) || 0;

      if (mlId && !bundleIds[mlId]) {
        if (!missing[mlId]) missing[mlId] = {nombre: nombre, qty: 0};
        missing[mlId].qty += qty;
      }
    }
  }

  // 결과를 묶음설정 시트 옆에 기록
  var missingKeys = Object.keys(missing);
  if (missingKeys.length === 0) {
    ui.alert("✅ 모든 ML_ID가 묶음설정에 있어요!");
    return;
  }

  // 묶음설정 시트에 누락 목록 추가 (H열부터)
  var headerRow = ["누락 ML_ID", "상품명", "오늘 판매수량", "SKU 입력필요", "묶음수량 입력필요"];
  bundleSheet.getRange(1, 7, 1, headerRow.length)
    .setValues([headerRow])
    .setBackground("#E74C3C").setFontColor("#FFFFFF").setFontWeight("bold");

  var missingRows = missingKeys.map(function(id) {
    return [id, missing[id].nombre, missing[id].qty, "← SKU 입력", "← 묶음수량 입력"];
  });

  bundleSheet.getRange(2, 7, missingRows.length, 5).setValues(missingRows);

  ui.alert(
    "⚠️ 묶음설정에 없는 ML_ID: " + missingKeys.length + "개\n\n" +
    "묶음설정 시트 G~K열에 목록을 기록했어요!\n\n" +
    "SKU와 묶음수량을 입력한 후\n" +
    "A~E열로 복사해서 추가해주세요."
  );
}

function autoRefreshMLTokenDaily() {
  var result = refreshMLToken("LUCAS");

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("ML토큰갱신로그");
  if (!logSheet) {
    logSheet = ss.insertSheet("ML토큰갱신로그");
    logSheet.getRange(1,1,1,2).setValues([["실행시간","상태"]])
      .setBackground("#1a2a3a").setFontColor("#FFFFFF").setFontWeight("bold");
  }

  var status = result ? "✅ 토큰 갱신 성공" : "❌ 토큰 갱신 실패";
  logSheet.appendRow([new Date(), status]);
  console.log("ML 토큰 자동갱신: " + status);
}

function manualRefreshMLToken() {
  var ui     = SpreadsheetApp.getUi();
  var result = refreshMLToken("LUCAS");

  if (result) {
    ui.alert("✅ 토큰 갱신 성공!\n\n내일부터 자동으로 갱신돼요.");
  } else {
    // 갱신 실패시 NAOS JK Trading 앱으로 새 토큰 발급 URL 안내
    ui.alert(
      "❌ 자동 갱신 실패\n\n" +
      "아래 URL에서 새 코드를 발급받아주세요\n" +
      "(Lucas 계정 로그인 상태에서):\n\n" +
      "https://auth.mercadolibre.cl/authorization" +
      "?response_type=code" +
      "&client_id=3465154046014198" +  // NAOS JK Trading 앱
      "&redirect_uri=https://www.google.com\n\n" +
      "TG-... 코드를 setNewMLToken 함수에 입력 후 실행하세요."
    );
  }
}

function checkTokenSettings() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = getMLTokenSheet();

  var clientId     = String(sheet.getRange(2, 2).getValue() || "").trim();
  var clientSecret = String(sheet.getRange(3, 2).getValue() || "").trim();
  var accessToken  = String(sheet.getRange(4, 2).getValue() || "").trim();
  var refreshToken = String(sheet.getRange(5, 2).getValue() || "").trim();

  ui.alert(
    "📋 ML토큰설정 현재값\n\n" +
    "B2 CLIENT_ID: "     + clientId     + "\n" +
    "B3 CLIENT_SECRET: " + clientSecret.substring(0,8) + "...\n" +
    "B4 ACCESS_TOKEN: "  + accessToken.substring(0,15)  + "...\n" +
    "B5 REFRESH_TOKEN: " + refreshToken.substring(0,15) + "...\n\n" +
    (clientId === "3465154046014198"
      ? "✅ NAOS JK Trading 앱 설정 완료"
      : "❌ CLIENT_ID가 잘못됨!\nB2셀을 3465154046014198로 수정해주세요")
  );
}

// ========================
// ML비교결과 실시간 업데이트
// ========================
function updateMLComparison() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

 var compSheet = ss.getSheetByName("🔍ML비교결과");
  if (!compSheet) {
    ui.alert("❌ ML비교결과 시트를 찾을 수 없어요!");
    return;
  }

  var lastRow = compSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("❌ 데이터가 없어요!");
    return;
  }

  var token = getMLToken("LUCAS");
  if (!token) {
    ui.alert("❌ 토큰 없음\n메뉴 → 토큰 수동 갱신을 실행해주세요.");
    return;
  }

  // 헤더에서 열 위치 파악
  var headers = compSheet.getRange(1, 1, 1, compSheet.getLastColumn()).getValues()[0];
  var colMap  = {};
  for (var h = 0; h < headers.length; h++) {
    colMap[String(headers[h]).trim()] = h + 1; // 1-based
  }

  var COL_SKU    = colMap["sku"]      || colMap["SKU"]      || 1;
  var COL_ML_ID  = colMap["ML_ID"]    || 4;
  var COL_STATUS = colMap["ML상태"]   || 6;
  var COL_STOCK  = colMap["ML재고"]   || 7;
  var COL_PRICE  = colMap["ML가격"]   || 8;
  var COL_NAOS   = colMap["NAOS재고"] || 9;
  var COL_REMARK = colMap["비고"]     || 11;
  var COL_URL    = colMap['URL']    || colMap['url']    || 12;

  // NAOS 재고에서 SKU → 재고수량 맵
  var invSheet = ss.getSheetByName(NOMBRE_INVENTARIO);
  var invData  = invSheet.getDataRange().getValues();
  var naosMap  = {};
  for (var j = 1; j < invData.length; j++) {
    var s = String(invData[j][1] || "").trim();
    if (s) naosMap[s] = Number(invData[j][6]) || 0;
  }

  // ML_ID 목록 수집
  var data   = compSheet.getRange(2, 1, lastRow-1, compSheet.getLastColumn()).getValues();
  var mlIds  = [];
  var rowMap = {}; // ML_ID → 행 번호들

  for (var i = 0; i < data.length; i++) {
    var mlId = String(data[i][COL_ML_ID-1] || "").trim();
    if (!mlId) continue;
    if (!rowMap[mlId]) rowMap[mlId] = [];
    rowMap[mlId].push(i + 2); // 실제 행 번호
    if (mlIds.indexOf(mlId) === -1) mlIds.push(mlId);
  }

  if (mlIds.length === 0) {
    ui.alert("❌ ML_ID가 없어요!\nD열에 ML_ID를 입력해주세요.");
    return;
  }

  // ML API로 상품 정보 일괄 조회 (20개씩 배치)
  var mlDataMap = {}; // ML_ID → {status, stock, price}
  var batchSize = 20;

  for (var b = 0; b < mlIds.length; b += batchSize) {
    var batch   = mlIds.slice(b, b + batchSize);
    var batchStr= batch.join(",");

    try {
      var resp = UrlFetchApp.fetch(
        "https://api.mercadolibre.com/items?ids=" + batchStr,
        {
          headers: {"Authorization": "Bearer " + token},
          muteHttpExceptions: true
        }
      );

      if (resp.getResponseCode() !== 200) {
        console.log("배치 조회 오류: " + resp.getResponseCode());
        continue;
      }

      var results = JSON.parse(resp.getContentText());
      for (var r = 0; r < results.length; r++) {
        var item = results[r].body || results[r];
        if (!item || !item.id) continue;

        mlDataMap[item.id] = {
          status: item.status       || "unknown",
          stock:  item.available_quantity !== undefined ? item.available_quantity : 0,
          price:  item.price        || 0,
          title:  item.title        || ""
        };
      }
    } catch(e) {
      console.log("배치 오류: " + e.message);
    }
  }

  // 시트 업데이트
  var updCount    = 0;
  var noIdCount   = 0;
  var errorCount  = 0;

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var sku    = String(data[i][COL_SKU-1]   || "").trim();
    var mlId   = String(data[i][COL_ML_ID-1] || "").trim();

    // NAOS 재고 업데이트
    if (sku && naosMap[sku] !== undefined) {
      compSheet.getRange(rowNum, COL_NAOS).setValue(naosMap[sku]);
    }

    // ML_ID 없으면 스킵
    if (!mlId) {
      noIdCount++;
      compSheet.getRange(rowNum, COL_STATUS).setValue("").setBackground(null);
      compSheet.getRange(rowNum, COL_REMARK)
        .setValue("ML미등록").setFontColor("#E74C3C");
      continue;
    }

    var mlInfo = mlDataMap[mlId];
    if (!mlInfo) {
      errorCount++;
      continue;
    }

    // ML상태 업데이트
    var statusCell = compSheet.getRange(rowNum, COL_STATUS);
    statusCell.setValue(mlInfo.status);
    if (mlInfo.status === "active") {
      statusCell.setBackground("#D5E8D4").setFontColor("#27AE60").setFontWeight("bold");
      statusCell.setValue("Activo");
    } else if (mlInfo.status === "paused") {
      statusCell.setBackground("#FFE6CC").setFontColor("#E67E22").setFontWeight("bold");
      statusCell.setValue("Pausado");
    } else {
      statusCell.setBackground("#F8CECC").setFontColor("#C0392B").setFontWeight("bold");
    }

    // ML재고 업데이트
    var stockCell = compSheet.getRange(rowNum, COL_STOCK);
    stockCell.setValue(mlInfo.stock);
    if (mlInfo.stock === 0) {
      stockCell.setBackground("#F8CECC").setFontColor("#C0392B").setFontWeight("bold");
    } else if (mlInfo.stock <= 5) {
      stockCell.setBackground("#FFE6CC").setFontColor("#E67E22").setFontWeight("bold");
    } else {
      stockCell.setBackground(null).setFontColor("#27AE60").setFontWeight("bold");
    }

    // ML가격 업데이트
    compSheet.getRange(rowNum, COL_PRICE)
      .setValue(mlInfo.price)
      .setNumberFormat("#,##0");

    // URL 업데이트
   var mlUrlVal = 'https://www.mercadolibre.cl/p/' + mlId;
    var COL_URL_COMP = compSheet.getLastColumn();
    // URL 열 찾기
    var hRow = compSheet.getRange(1,1,1,compSheet.getLastColumn()).getValues()[0];
    for (var uh = 0; uh < hRow.length; uh++) {
      if (String(hRow[uh]).trim().toUpperCase() === 'URL') {
        COL_URL_COMP = uh + 1; break;
      }
    }
    compSheet.getRange(rowNum, COL_URL_COMP)
      .setValue(mlUrlVal)
      .setFontColor('#2980B9')
      .setHorizontalAlignment('center');

    // 비고 업데이트
    var naosStock = naosMap[sku] !== undefined ? naosMap[sku] : -1;
    var remark    = "";
    var remarkBg  = null;
    var remarkFg  = "#333333";

    if (mlInfo.status === "active" || mlInfo.status === "Activo") {
      if (naosStock <= 0) {
        remark   = "⚠️ NAOS재고없음";
        remarkBg = "#FFE6CC";
        remarkFg = "#E67E22";
      } else if (mlInfo.stock <= 0) {
        remark   = "⚠️ ML재고없음";
        remarkBg = "#FFE6CC";
        remarkFg = "#E67E22";
      } else {
        remark   = "✅ 정상";
        remarkBg = "#D5E8D4";
        remarkFg = "#27AE60";
      }
    } else {
      remark   = "⏸ 판매중지";
      remarkBg = "#F8CECC";
      remarkFg = "#C0392B";
    }

    compSheet.getRange(rowNum, COL_REMARK)
      .setValue(remark)
      .setBackground(remarkBg)
      .setFontColor(remarkFg)
      .setFontWeight("bold");

    updCount++;
  }

  // 마지막 업데이트 시간 기록
  compSheet.getRange(1, compSheet.getLastColumn() + 1)
    .setValue("최종업데이트: " +
      Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm:ss"))
    .setFontColor("#888888").setFontSize(8);

  ui.alert(
    "✅ ML비교결과 업데이트 완료!\n\n" +
    "업데이트: " + updCount   + "개\n" +
    "ML미등록: " + noIdCount  + "개\n" +
    "오류:     " + errorCount + "개"
  );
}

// 자동 업데이트 (트리거용)
function autoUpdateMLComparison() {
  try {
    updateMLComparison();
  } catch(e) {
    console.log("ML비교결과 자동업데이트 오류: " + e.message);
  }
}

function checkSheetNames() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var names  = sheets.map(function(s){ return '"' + s.getName() + '"'; });
  SpreadsheetApp.getUi().alert(
    "시트 목록:\n\n" + names.join("\n")
  );
}

function testMLListingPrice() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = getMLTokenSheet();
  var token = String(sheet.getRange(4, 2).getValue() || "").trim();

  // MLC1571522135 상품 정보 먼저 가져오기
  var itemResp = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/items?ids=MLC1571522135",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  var itemData = JSON.parse(itemResp.getContentText());
  var item     = itemData[0] ? (itemData[0].body || itemData[0]) : {};

  var price      = item.price        || 0;
  var categoryId = item.category_id  || "";
  var title      = item.title        || "";

  var msg = "상품 정보\n"
    + "제목: "      + title.substring(0,30) + "\n"
    + "가격: $"     + price + "\n"
    + "카테고리ID: " + categoryId + "\n\n";

  // 수수료 조회
  if (price && categoryId) {
    var feeResp = UrlFetchApp.fetch(
      "https://api.mercadolibre.com/sites/MLC/listing_prices" +
      "?price=" + price +
      "&category_id=" + categoryId +
      "&listing_type_id=gold_special",
      {muteHttpExceptions:true}
    );
    msg += "수수료 API 응답코드: " + feeResp.getResponseCode() + "\n";
    if (feeResp.getResponseCode() === 200) {
      var feeData   = JSON.parse(feeResp.getContentText());
      var salesFee  = feeData.sale_fee_amount       || 0;
      var listFee   = feeData.listing_fee_amount     || 0;
      var shipping  = feeData.shipping_fee_amount    || 0;
      msg += "판매수수료: $"  + salesFee  + "\n"
           + "등록수수료: $"  + listFee   + "\n"
           + "배송수수료: $"  + shipping  + "\n"
           + "실수령액: $"    + (price - salesFee - listFee) + "\n";
    } else {
      msg += feeResp.getContentText().substring(0,150) + "\n";
    }
  }

  ui.alert(msg);
}
function testMLPermalink() {
  var ui    = SpreadsheetApp.getUi();
  var token = getMLTokenSheet().getRange(4,2).getValue();
  var mlId  = "MLC1580443025"; // 아까 테스트한 제품

  var resp = UrlFetchApp.fetch(
    'https://api.mercadolibre.com/items/' + mlId,
    {headers:{'Authorization':'Bearer '+token}, muteHttpExceptions:true}
  );

  var item = JSON.parse(resp.getContentText());

  ui.alert(
    '응답코드: ' + resp.getResponseCode() + '\n\n' +
    'permalink: ' + (item.permalink||'없음') + '\n\n' +
    'catalog_product_id: ' + (item.catalog_product_id||'없음') + '\n\n' +
    '올바른 URL:\n' +
    'https://www.mercadolibre.cl/tirador-manilla-eroom-acero-inoxidable-negro-mate64x-100mm-10uni-puerta-cajon/p/MLC46393840?pdp_filters=item_id:MLC1580443025'
  );
}

function restoreMLUrls() {
  var ui        = SpreadsheetApp.getUi();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var compSheet = ss.getSheetByName('🔍ML비교결과');
  if (!compSheet) { ui.alert('시트 없음'); return; }

  var headers = compSheet.getRange(1,1,1,compSheet.getLastColumn()).getValues()[0];
  var colMap  = {};
  for (var h = 0; h < headers.length; h++) {
    colMap[String(headers[h]).trim()] = h+1;
  }

  var COL_ML_ID = colMap['ML_ID'] || 4;
  var COL_URL   = colMap['URL']   || 12;

  var lastRow = compSheet.getLastRow();
  var data    = compSheet.getRange(2,1,lastRow-1,compSheet.getLastColumn()).getValues();
  var token   = getMLTokenSheet().getRange(4,2).getValue();
  var updates = 0;

  for (var i = 0; i < data.length; i++) {
    var mlId = String(data[i][COL_ML_ID-1]||'').trim();
    if (!mlId) continue;

    try {
      var resp = UrlFetchApp.fetch(
        'https://api.mercadolibre.com/items/' + mlId,
        {headers:{'Authorization':'Bearer '+token}, muteHttpExceptions:true}
      );
      if (resp.getResponseCode() !== 200) continue;

      var item      = JSON.parse(resp.getContentText());

      // catalog_product_id로 URL 생성
      var catId     = item.catalog_product_id || '';
      var permalink = '';

      if (catId) {
        // 카탈로그 상품 URL
        permalink = 'https://www.mercadolibre.cl/p/' + catId;
      } else {
        // 카탈로그 없는 상품은 ML_ID로 직접 링크
        permalink = 'https://www.mercadolibre.cl/p/' + mlId;
      }

      compSheet.getRange(i+2, COL_URL)
        .setValue(permalink)
        .setFontColor('#2980B9')
        .setHorizontalAlignment('center');
      updates++;
      Utilities.sleep(100);

    } catch(e) {
      console.log('URL 복구 오류 ' + mlId + ': ' + e.message);
    }
  }

  ui.alert('✅ URL 복구 완료!\n\n' + updates + '개 URL 업데이트됨');
}

function testMLCompHeaders() {
  var ui        = SpreadsheetApp.getUi();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var compSheet = ss.getSheetByName('🔍ML비교결과');
  if (!compSheet) { ui.alert('시트 없음'); return; }

  var headers  = compSheet.getRange(1,1,1,compSheet.getLastColumn()).getValues()[0];
  var lastData = compSheet.getRange(2,1,1,compSheet.getLastColumn()).getValues()[0];

  var msg = '헤더 확인\n\n';
  for (var i = 0; i < headers.length; i++) {
    msg += (i+1)+'열('+String.fromCharCode(65+i)+'): '
      + headers[i] + ' → ' + String(lastData[i]||'').substring(0,20) + '\n';
  }
  ui.alert(msg);
}

function restoreMLUrls() {
  var ui        = SpreadsheetApp.getUi();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var compSheet = ss.getSheetByName('🔍ML비교결과');
  if (!compSheet) { ui.alert('시트 없음'); return; }

  var lastRow = compSheet.getLastRow();
  if (lastRow < 2) { ui.alert('데이터 없음'); return; }

  var mlIdData = compSheet.getRange(2, 4, lastRow-1, 1).getValues();
  var newUrls  = [];
  var updates  = 0;

  for (var i = 0; i < mlIdData.length; i++) {
    var mlId = String(mlIdData[i][0]||'').trim();
    if (!mlId) {
      newUrls.push(['']);
      continue;
    }
    // articulo 형식으로 생성 (판매자 게시물 직접 링크)
    // MLC1401273421 → MLC-1401273421
    var numPart = mlId.replace('MLC','');
    var newUrl  = 'https://articulo.mercadolibre.cl/MLC-' + numPart;
    newUrls.push([newUrl]);
    updates++;
  }

  compSheet.getRange(2, 12, lastRow-1, 1)
    .setValues(newUrls)
    .setFontColor('#2980B9')
    .setHorizontalAlignment('center');

  ui.alert('✅ URL 복구 완료!\n\n' + updates + '개 URL 생성됨');
}

function testMLOrderUrl() {
  var ui    = SpreadsheetApp.getUi();
  var token = getMLToken("LUCAS");

  // 방법 1: 날짜 없이
  var r1 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=323449286&order.status=paid&sort=date_desc&limit=5",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  var msg = "방법1 (날짜없음): " + r1.getResponseCode() + "\n";

  // 방법 2: 날짜 포함
  var dateFrom = new Date();
  dateFrom.setDate(dateFrom.getDate() - 7);
  var r2 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=323449286&order.status=paid&order.date_created.from="+dateFrom.toISOString()+"&sort=date_desc&limit=5",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법2 (7일): " + r2.getResponseCode() + "\n";

  // 방법 3: offset 포함
  var r3 = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search?seller=323449286&order.status=paid&sort=date_desc&limit=5&offset=0",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );
  msg += "방법3 (offset): " + r3.getResponseCode() + "\n";

  if (r1.getResponseCode() === 200) {
    var d = JSON.parse(r1.getContentText());
    msg += "\n방법1 성공!\n주문수: " + (d.paging?d.paging.total:0) + "건";
  } else {
    msg += "\n방법1 오류내용: " + r1.getContentText().substring(0,100);
  }

  ui.alert(msg);
}

function testMLFlexType() {
  var ui    = SpreadsheetApp.getUi();
  var token = getMLToken("LUCAS");

  var dateFrom = new Date();
  dateFrom.setDate(dateFrom.getDate() - 7);

  var resp = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search" +
    "?seller=323449286&order.status=paid" +
    "&order.date_created.from=" + dateFrom.toISOString() +
    "&sort=date_desc&limit=10",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );

  var orders = JSON.parse(resp.getContentText()).results || [];
  var msg    = "주문별 logistic_type 확인\n\n";

  for (var o = 0; o < orders.length; o++) {
    var order    = orders[o];
    var shipping = order.shipping || {};
    var logistic = shipping.logistic_type || "없음";
    var tags     = (order.tags || []).join(", ");

    msg += "주문" + (o+1) + ":\n";
    msg += "  logistic_type: " + logistic + "\n";
    msg += "  tags: " + tags + "\n\n";
  }

  ui.alert(msg);
}

function testMLFlexType2() {
  var ui    = SpreadsheetApp.getUi();
  var token = getMLToken("LUCAS");

  var dateFrom = new Date();
  dateFrom.setDate(dateFrom.getDate() - 7);

  var resp = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/orders/search" +
    "?seller=323449286&order.status=paid" +
    "&order.date_created.from=" + dateFrom.toISOString() +
    "&sort=date_desc&limit=5",
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );

  var orders = JSON.parse(resp.getContentText()).results || [];
  var msg    = "shipping 전체 확인\n\n";

  for (var o = 0; o < Math.min(3, orders.length); o++) {
    var order    = orders[o];
    var shipping = order.shipping || {};
    msg += "주문" + (o+1) + ":\n";
    msg += JSON.stringify(shipping).substring(0,200) + "\n\n";
  }

  ui.alert(msg);
}

function testMLShippingDetail() {
  var ui    = SpreadsheetApp.getUi();
  var token = getMLToken("LUCAS");

  // 주문1의 shipping ID로 상세 조회
  var shippingId = "47089154840";

  var resp = UrlFetchApp.fetch(
    "https://api.mercadolibre.com/shipments/" + shippingId,
    {headers:{"Authorization":"Bearer "+token}, muteHttpExceptions:true}
  );

  var msg = "응답코드: " + resp.getResponseCode() + "\n\n";
  if (resp.getResponseCode() === 200) {
    var data = JSON.parse(resp.getContentText());
    msg += "logistic_type: " + (data.logistic_type||"없음") + "\n";
    msg += "shipping_mode: " + (data.shipping_mode||"없음") + "\n";
    msg += "service_id: "    + (data.service_id||"없음") + "\n";
    msg += "tags: " + JSON.stringify(data.tags||[]) + "\n";
  } else {
    msg += resp.getContentText().substring(0,200);
  }

  ui.alert(msg);
}

function addCheckboxToOldRows() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ML판매내역");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // A열 전체 체크박스 확인 후 없는 행에 추가
  var aVals = sheet.getRange(2,1,lastRow-1,1).getValues();
  for (var i = 0; i < aVals.length; i++) {
    var val = aVals[i][0];
    // 체크박스가 아닌 경우 추가
    if (val !== true && val !== false) {
      sheet.getRange(i+2, 1).insertCheckboxes();
    }
  }

  // 완료주문번호에 2~7행 주문ID 추가
  var doneSheet = ss.getSheetByName("✅완료주문번호");
  var data      = sheet.getRange(2,1,lastRow-1,15).getValues();
  var newIds    = [];

  // 기존 완료주문번호 확인
  var doneIds = {};
  if (doneSheet && doneSheet.getLastRow() >= 2) {
    var doneData = doneSheet.getRange(2,1,doneSheet.getLastRow()-1,1).getValues();
    for (var d = 0; d < doneData.length; d++) {
      doneIds[String(doneData[d][0]).trim()] = true;
    }
  }

  // C열 주문ID 읽기
  for (var i = 0; i < data.length; i++) {
    var orderId = String(data[i][2]||"").trim();
    if (orderId && !doneIds[orderId]) {
      newIds.push([orderId]);
      doneIds[orderId] = true;
    }
  }

  // 완료주문번호에 추가
  if (newIds.length > 0 && doneSheet) {
    var lastDoneRow = doneSheet.getLastRow() + 1;
    doneSheet.getRange(lastDoneRow,1,newIds.length,1)
      .setValues(newIds).setNumberFormat("@");
  }

  SpreadsheetApp.getUi().alert(
    "✅ 완료!\n\n" +
    "체크박스 추가됨\n" +
    "완료주문번호 추가: " + newIds.length + "건"
  );
}

function testVentasSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var msg = "NOMBRE_VENTAS: " + NOMBRE_VENTAS + "\n\n";

  var ventasSheet = ss.getSheetByName(NOMBRE_VENTAS);
  if (ventasSheet) {
    msg += "✅ Ventas 시트 찾음!\n";
    msg += "마지막 행: " + ventasSheet.getLastRow() + "\n";
    msg += "열 수: " + ventasSheet.getLastColumn();
  } else {
    msg += "❌ Ventas 시트 없음!\n\n";
    msg += "현재 시트 목록:\n";
    ss.getSheets().forEach(function(s) {
      msg += "- " + s.getName() + "\n";
    });
  }

  ui.alert(msg);
}

function testVentasStructure() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_VENTAS);

  // 1행 헤더 확인
  var headers  = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  // 2행 데이터 확인
  var row2     = sheet.getRange(2,1,1,sheet.getLastColumn()).getValues()[0];

  var msg = "Ventas 시트 구조\n\n";
  msg += "총 열수: " + sheet.getLastColumn() + "\n\n";
  msg += "헤더:\n";
  for (var i = 0; i < headers.length; i++) {
    msg += (i+1)+"열: " + headers[i] + " → " + String(row2[i]||"").substring(0,15) + "\n";
  }
  ui.alert(msg);
}

// ============================================================
// ML 토큰 자동 갱신 + 판매내역 자동 가져오기
// ============================================================

// 트리거 전체 재설정 (중복 제거 + 올바르게 설정)
function setupAllTriggers() {
  var ui = SpreadsheetApp.getUi();

  // 기존 트리거 전체 삭제
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // ① ML 토큰 갱신 — 매일 4회 (0시/7시/13시/19시)
  ScriptApp.newTrigger('autoRefreshMLTokenDaily')
    .timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger('autoRefreshMLTokenDaily')
    .timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('autoRefreshMLTokenDaily')
    .timeBased().atHour(13).everyDays(1).create();
  ScriptApp.newTrigger('autoRefreshMLTokenDaily')
    .timeBased().atHour(19).everyDays(1).create();

  // ② ML 판매내역 자동 가져오기 — 30분마다
  ScriptApp.newTrigger('autoFetchMLSales')
    .timeBased().everyMinutes(30).create();

  // ③ NAOS 재고 자동 업데이트 — 1시간마다
  ScriptApp.newTrigger('autoUpdateNaosStock')
    .timeBased().everyHours(1).create();

  ui.alert(
    '✅ 트리거 설정 완료!\n\n' +
    '① ML 토큰 갱신: 매일 4회 (0/7/13/19시)\n' +
    '② ML 판매내역: 30분마다 자동\n' +
    '③ NAOS 재고 업데이트: 1시간마다'
  );
}

// ML 판매내역 자동 가져오기
function autoFetchMLSales() {
  try {
    // 토큰 확인
    var ss         = SpreadsheetApp.openById('1hw8M2WsmmTY-Obz12N3WUBi7Fv1tlrs9tGBUGvwHWrk');
    var propToken  = PropertiesService.getScriptProperties().getProperty('ML_ACCESS_TOKEN');
    var sheetToken = '';

    // 토큰 시트에서도 확인
    var tokenSheet = ss.getSheetByName('ML토큰');
    if (tokenSheet) {
      sheetToken = tokenSheet.getRange(1,2).getValue();
    }

    var token = propToken || sheetToken;

    if (!token) {
      // 토큰 없으면 먼저 갱신 시도
      autoRefreshMLTokenDaily();
      token = PropertiesService.getScriptProperties().getProperty('ML_ACCESS_TOKEN');
    }

    if (!token) {
      console.log('autoFetchMLSales: 토큰 없음 - 건너뜀');
      return;
    }

    // 판매내역 가져오기
    fetchMLSalesAndDeduct();
    console.log('autoFetchMLSales: 완료 ' + new Date().toISOString());

  // 저장 후 날짜순 자동 정렬
  sortMLSalesByDate();

} catch(e) {
  console.log('fetchMLSalesAndDeduct 오류: ' + e.message);
}
}

// 토큰 상태 확인
function checkMLTokenStatus() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.openById('1hw8M2WsmmTY-Obz12N3WUBi7Fv1tlrs9tGBUGvwHWrk');
  var prop = PropertiesService.getScriptProperties();

  var accessToken  = prop.getProperty('ML_ACCESS_TOKEN')  || '';
  var refreshToken = prop.getProperty('ML_REFRESH_TOKEN') || '';
  var clientId     = prop.getProperty('ML_CLIENT_ID')     || '';

  // 토큰 시트 확인
  var tokenSheet = ss.getSheetByName('ML토큰');
  var sheetData  = '';
  if (tokenSheet) {
    var vals = tokenSheet.getRange(1,1,5,2).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0]) sheetData += vals[i][0] + ': ' + String(vals[i][1]).substring(0,20) + '\n';
    }
  }

  // 트리거 목록
  var triggers  = ScriptApp.getProjectTriggers();
  var triggerList = triggers.map(function(t) {
    return t.getHandlerFunction();
  }).join('\n');

  ui.alert(
    'ML 토큰 상태\n\n' +
    'PropertiesService:\n' +
    'ACCESS_TOKEN: '  + (accessToken  ? accessToken.substring(0,20)  + '...' : '없음') + '\n' +
    'REFRESH_TOKEN: ' + (refreshToken ? refreshToken.substring(0,20) + '...' : '없음') + '\n' +
    'CLIENT_ID: '     + (clientId     ? clientId.substring(0,10)     + '...' : '없음') + '\n\n' +
    'ML토큰 시트:\n' + (sheetData || '없음') + '\n\n' +
    '트리거 목록:\n' + (triggerList || '없음')
  );
}

function saveMLSecret() {
  PropertiesService.getScriptProperties()
    .setProperty('ML_CLIENT_SECRET', 'sCj1L5Q6mz7iwNzArwo2IcuZo4nXTpoq');
  SpreadsheetApp.getUi().alert('✅ 저장완료!');
}

function getMLAuthUrl() {
  var ui       = SpreadsheetApp.getUi();
  var clientId = '3465154046014198';
  var redirect = 'https://www.google.com';
  var url      = 'https://auth.mercadolibre.cl/authorization'
    + '?response_type=code'
    + '&client_id=' + clientId
    + '&redirect_uri=' + encodeURIComponent(redirect);

  ui.alert(
    '아래 URL을 브라우저에서 열어주세요:\n\n' +
    url + '\n\n' +
    '로그인 후 google.com으로 이동되면\n' +
    '브라우저 주소창에서\n' +
    '"code=TG-XXXXX" 부분을 복사해주세요!'
  );
}

function getMLAppInfo() {
  var ui       = SpreadsheetApp.getUi();
  var clientId = '3465154046014198';

  var response = UrlFetchApp.fetch(
    'https://api.mercadolibre.com/applications/' + clientId,
    {muteHttpExceptions: true}
  );

  var data = JSON.parse(response.getContentText());
  ui.alert(
    'ML 앱 정보:\n\n' +
    '이름: '          + (data.name        || '없음') + '\n' +
    'Redirect URI: '  + (data.redirect_uri || '없음') + '\n' +
    'Site ID: '       + (data.site_id      || '없음') + '\n\n' +
    '전체:\n' + JSON.stringify(data).substring(0, 400)
  );
}

function getTokenFromCode() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    '인증 완료 URL 입력',
    'https://www.google.com/?code=TG-6a14b25b912224000143eec0-323449286',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var input = resp.getResponseText().trim();

  // URL에서 code 자동 추출
  var code = '';
  var match = input.match(/[?&]code=([^&]+)/);
  if (match) {
    code = match[1];
  } else {
    // URL이 아니라 code만 입력한 경우
    code = input;
  }

  code = code.trim();

  if (!code) {
    ui.alert('❌ code를 찾을 수 없어요!\nURL을 다시 확인해주세요.');
    return;
  }

  var prop   = PropertiesService.getScriptProperties();
  var secret = prop.getProperty('ML_CLIENT_SECRET');

  if (!secret) {
    ui.alert('❌ Client Secret 없음!\nsaveMLSecret() 먼저 실행해주세요!');
    return;
  }

  ui.alert('추출된 code:\n' + code + '\n\n이 code로 토큰 발급합니다...');

  var response = UrlFetchApp.fetch(
    'https://api.mercadolibre.com/oauth/token', {
    method : 'POST',
    headers: {'Content-Type':'application/x-www-form-urlencoded'},
    payload: 'grant_type=authorization_code'
      + '&client_id=3465154046014198'
      + '&client_secret=' + secret
      + '&code=' + code
      + '&redirect_uri=' + encodeURIComponent('https://www.google.com'),
    muteHttpExceptions: true
  });

  var data = JSON.parse(response.getContentText());

  if (data.access_token) {
    prop.setProperty('ML_ACCESS_TOKEN',  data.access_token);
    prop.setProperty('ML_REFRESH_TOKEN', data.refresh_token);
    prop.setProperty('ML_CLIENT_ID',     '3465154046014198');

    // ML토큰 시트에도 저장
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ML토큰');
    if (!sheet) sheet = ss.insertSheet('ML토큰');
    sheet.clearContents();
    sheet.getRange(1,1,4,2).setValues([
      ['access_token',  data.access_token],
      ['refresh_token', data.refresh_token],
      ['client_id',     '3465154046014198'],
      ['updated_at',    new Date().toISOString()]
    ]);

    ui.alert(
      '✅ 토큰 저장 완료!\n\n' +
      'ACCESS_TOKEN: ' + data.access_token.substring(0,25) + '...\n' +
      'REFRESH_TOKEN: ' + data.refresh_token.substring(0,25) + '...\n\n' +
      '다음: autoFetchMLSales() 실행!'
    );
  } else {
    ui.alert('❌ 오류:\n' + JSON.stringify(data));
  }
}

function sortMLSalesByDate() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ML판매내역');
    if (!sheet || sheet.getLastRow() < 3) return;

    var startRow = 2;
    var lastRow  = sheet.getLastRow();
    var lastCol  = sheet.getLastColumn();
    var numRows  = lastRow - startRow + 1;
    if (numRows < 2) return;

    // B열 (처리일시) 기준 오름차순 정렬
    var range = sheet.getRange(startRow, 1, numRows, lastCol);
    range.sort({column: 2, ascending: true});

    console.log('sortMLSalesByDate 완료: ' + numRows + '행');
  } catch(e) {
    console.log('sortMLSalesByDate 오류: ' + e.message);
  }
}

function checkMLSalesSheet() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ML판매내역');

  if (!sheet) {
    ui.alert('ML판매내역 시트 없음!');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // 1~3행 데이터 확인
  var top3 = sheet.getRange(1,1,3,lastCol).getValues();
  var msg  = '시트 정보:\n';
  msg += '총 행수: ' + lastRow + '\n';
  msg += '총 열수: ' + lastCol + '\n\n';
  msg += '1행: ' + JSON.stringify(top3[0]).substring(0,80) + '\n';
  msg += '2행: ' + JSON.stringify(top3[1]).substring(0,80) + '\n';
  msg += '3행: ' + JSON.stringify(top3[2]).substring(0,80) + '\n\n';

  // 시트 보호 여부 확인
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  msg += '범위 보호: ' + protections.length + '개\n';

  var sheetProt = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  msg += '시트 보호: ' + sheetProt.length + '개\n';

  ui.alert(msg);
}

function exportImageURLs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventario");
  const lastRow = sheet.getLastRow();
  
  // Foto 컬럼 = H열 (8번째)
  const results = [];
  for (let i = 2; i <= lastRow; i++) {
    const sku  = sheet.getRange(i, 2).getValue();
    const name = sheet.getRange(i, 3).getValue();
    const foto = sheet.getRange(i, 8).getFormula(); // =IMAGE("URL") 수식
    if (foto && foto.includes("IMAGE")) {
      const url = foto.match(/IMAGE\("([^"]+)"/)?.[1] || "";
      results.push({ sku, name, url });
    }
  }
  
  Logger.log(JSON.stringify(results.slice(0, 5)));
  // 결과를 새 시트에 저장
  const out = ss.getSheetByName("ImageURLs") || ss.insertSheet("ImageURLs");
  out.clearContents();
  out.appendRow(["SKU", "Nombre", "Image URL"]);
  results.forEach(r => out.appendRow([r.sku, r.name, r.url]));
  SpreadsheetApp.getUi().alert(`완료: ${results.length}개 URL 추출됨`);
}

/**
 * 실행 함수: makePhotoFile()
 * - 현재 스프레드시트(NAOS재고)의 ImageURLs 시트 읽기
 * - 완전히 새로운 별도 구글 스프레드시트 파일 생성
 * - 파일명: 제품 사진.url
 */

/**
 * makePhotoFile() - 기존 버전 교체
 * 저장 파일명(F열), 드라이브 저장 URL(G열) 컬럼 추가 버전
 */
function makePhotoFile() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName("ImageURLs");

  if (!srcSheet) {
    const allSheets = ss.getSheets().map(s => s.getName());
    SpreadsheetApp.getUi().alert(
      "❌ ImageURLs 시트를 찾을 수 없습니다.\n\n" +
      "현재 파일의 시트 목록:\n" + allSheets.join("\n")
    );
    return;
  }

  const srcData  = srcSheet.getDataRange().getValues();
  const dataRows = srcData.slice(1).filter(row => row[0] && row[2]);

  if (dataRows.length === 0) {
    SpreadsheetApp.getUi().alert("❌ ImageURLs 시트에 데이터가 없습니다.");
    return;
  }

  // 새 스프레드시트 파일 생성
  const newSS    = SpreadsheetApp.create("제품 사진.url");
  const newSheet = newSS.getActiveSheet();
  newSheet.setName("제품 사진.url");

  // 헤더 (F열: 저장 파일명, G열: 드라이브 저장 URL 추가)
  const headers = [
    "No",                   // A
    "SKU",                  // B
    "제품명",               // C
    "현재 사진",            // D
    "✅ 새 사진 URL",       // E ← 여기에 URL 붙여넣기
    "저장 파일명",          // F ← SKU 기반 자동생성
    "드라이브 저장 URL",    // G ← savePhotosToFolder() 실행 후 자동입력
    "1688",                 // H
    "Mercado Libre",        // I
    "AliExpress",           // J
    "Alibaba",              // K
    "Google 이미지",        // L
    "완료"                  // M
  ];

  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 스타일
  const hRange = newSheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground("#1a73e8");
  hRange.setFontColor("#ffffff");
  hRange.setFontWeight("bold");
  hRange.setHorizontalAlignment("center");

  // 데이터 생성
  const outputRows = [];

  dataRows.forEach((row, idx) => {
    const sku      = String(row[0]).trim();
    const nombre   = String(row[1]).trim();
    const imageUrl = String(row[2]).trim();
    const q        = encodeURIComponent(nombre);

    // 저장 파일명: SKU 공백 → _ 로 변환
    const fileName = sku.replace(/\s+/g, "_");

    outputRows.push([
      idx + 1,          // A: No
      sku,              // B: SKU
      nombre,           // C: 제품명
      imageUrl,         // D: 현재 사진 (수식으로 교체 예정)
      "",               // E: 새 사진 URL ← 붙여넣기
      fileName,         // F: 저장 파일명 (자동생성)
      "",               // G: 드라이브 저장 URL (스크립트 실행 후 자동)
      "https://s.1688.com/selloffer/offer_search.htm?keywords=" + encodeURIComponent(nombre),
      "https://listado.mercadolibre.cl/" + q,
      "https://www.aliexpress.com/wholesale?SearchText=" + q,
      "https://www.alibaba.com/trade/search?SearchText=" + q,
      "https://www.google.com/search?q=" + q + "&tbm=isch",
      false             // M: 완료 체크박스
    ]);
  });

  // 일괄 입력
  newSheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);

  // 열 너비
  newSheet.setColumnWidth(1, 50);   // A: No
  newSheet.setColumnWidth(2, 130);  // B: SKU
  newSheet.setColumnWidth(3, 240);  // C: 제품명
  newSheet.setColumnWidth(4, 80);   // D: 현재 사진
  newSheet.setColumnWidth(5, 220);  // E: 새 사진 URL
  newSheet.setColumnWidth(6, 160);  // F: 저장 파일명
  newSheet.setColumnWidth(7, 180);  // G: 드라이브 저장 URL
  newSheet.setColumnWidth(8, 110);  // H: 1688
  newSheet.setColumnWidth(9, 130);  // I: ML
  newSheet.setColumnWidth(10, 120); // J: AliExpress
  newSheet.setColumnWidth(11, 110); // K: Alibaba
  newSheet.setColumnWidth(12, 130); // L: Google
  newSheet.setColumnWidth(13, 55);  // M: 완료

  // E열(새 사진 URL) 배경 연초록
  newSheet.getRange(2, 5, outputRows.length, 1).setBackground("#e6f4ea");

  // F열(저장 파일명) 배경 연노랑
  newSheet.getRange(2, 6, outputRows.length, 1).setBackground("#fff2cc");

  // 검색 URL 열 하이퍼링크 스타일
  [8, 9, 10, 11, 12].forEach(col => {
    newSheet.getRange(2, col, outputRows.length, 1)
      .setFontColor("#1155cc")
      .setTextStyle(SpreadsheetApp.newTextStyle().setUnderline(true).build());
  });

  // 완료 체크박스
  newSheet.getRange(2, 13, outputRows.length, 1).insertCheckboxes();

  // 완료 시 녹색 조건부 서식
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$M2=TRUE")
    .setBackground("#d9ead3")
    .setRanges([newSheet.getRange(2, 1, outputRows.length, headers.length)])
    .build();
  newSheet.setConditionalFormatRules([rule]);

  // 행 높이 + 현재 사진 미리보기
  newSheet.setRowHeightsForced(2, outputRows.length, 80);
  for (let i = 0; i < outputRows.length; i++) {
    const url = outputRows[i][3];
    if (url && url.startsWith("http")) {
      newSheet.getRange(i + 2, 4).setFormula('=IMAGE("' + url + '",4,70,70)');
    }
  }

  // 고정 행/열
  newSheet.setFrozenRows(1);
  newSheet.setFrozenColumns(3);

  // 완료 알림
  const newUrl = newSS.getUrl();
  SpreadsheetApp.getUi().alert(
    "✅ 새 파일 생성 완료!\n\n" +
    "파일명: 제품 사진.url\n" +
    "제품 수: " + outputRows.length + "개\n\n" +
    "파일 URL:\n" + newUrl + "\n\n" +
    "다음 단계:\n" +
    "1. E열(초록)에 고해상도 사진 URL 붙여넣기\n" +
    "2. savePhotosToFolder() 실행\n" +
    "   → 드라이브에 SKU명으로 자동 저장됨"
  );
}


/**
 * savePhotosToFolder()
 * E열(새 사진 URL)의 이미지를 드라이브 폴더에 SKU명으로 자동 저장
 *
 * 실행 전:
 *   아래 PHOTO_FILE_ID 에 "제품 사진.url" 파일 ID 입력
 *   (파일 열기 → 주소창 URL → /d/ 뒤 문자열)
 */
function savePhotosToFolder() {

  // ▼ "제품 사진.url" 파일 ID 입력
  const PHOTO_FILE_ID = "여기에_파일ID_입력";

  // ▼ 드라이브 저장 폴더명
  const FOLDER_NAME = "NAOS 고해상도 사진";

  if (PHOTO_FILE_ID === "여기에_파일ID_입력") {
    SpreadsheetApp.getUi().alert(
      "❌ PHOTO_FILE_ID를 입력해 주세요.\n\n" +
      "'제품 사진.url' 파일 URL에서\n" +
      "/d/ 뒤 문자열을 복사해 넣으세요.\n\n" +
      "예) https://docs.google.com/.../d/★여기★/edit"
    );
    return;
  }

  // 드라이브 폴더 찾기 또는 생성
  let folder;
  const folderIter = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folderIter.hasNext()) {
    folder = folderIter.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
  }

  // 시트 데이터 읽기
  const photoSS    = SpreadsheetApp.openById(PHOTO_FILE_ID);
  const photoSheet = photoSS.getSheetByName("제품 사진.url");
  if (!photoSheet) {
    SpreadsheetApp.getUi().alert("❌ '제품 사진.url' 시트를 찾을 수 없습니다.");
    return;
  }

  const data = photoSheet.getDataRange().getValues();
  // 컬럼: No(0) SKU(1) 제품명(2) 현재사진(3) 새URL(4) 파일명(5) 드라이브URL(6) ...

  let saved   = 0;
  let skipped = 0;
  let errors  = 0;

  for (let i = 1; i < data.length; i++) {
    const newUrl   = String(data[i][4]).trim(); // E열: 새 사진 URL
    const fileName = String(data[i][5]).trim(); // F열: 저장 파일명

    // URL 없으면 스킵
    if (!newUrl || !newUrl.startsWith("http")) {
      skipped++;
      continue;
    }

    // 이미 저장된 경우 스킵 (중복 방지)
    if (data[i][6] && String(data[i][6]).startsWith("http")) {
      skipped++;
      continue;
    }

    try {
      // 이미지 다운로드
      const response    = UrlFetchApp.fetch(newUrl, {muteHttpExceptions: true});
      const blob        = response.getBlob();
      const contentType = blob.getContentType() || "image/jpeg";

      // 확장자 자동 감지
      let ext = ".jpg";
      if (contentType.includes("png"))  ext = ".png";
      if (contentType.includes("webp")) ext = ".webp";
      if (contentType.includes("gif"))  ext = ".gif";

      // 드라이브 폴더에 SKU명으로 저장
      blob.setName(fileName + ext);
      const savedFile = folder.createFile(blob);

      // G열(드라이브 저장 URL) 자동 입력
      photoSheet.getRange(i + 1, 7).setValue(savedFile.getUrl());

      // M열(완료) 체크
      photoSheet.getRange(i + 1, 13).setValue(true);

      saved++;

      // 50개마다 2초 대기 (할당량 초과 방지)
      if (saved % 50 === 0) Utilities.sleep(2000);

    } catch(e) {
      Logger.log("오류 [" + fileName + "]: " + e.toString());
      errors++;
    }
  }

  SpreadsheetApp.getUi().alert(
    "✅ 드라이브 저장 완료!\n\n" +
    "저장된 사진: " + saved + "개\n" +
    "스킵: " + skipped + "개\n" +
    "오류: " + errors + "개\n\n" +
    "저장 폴더: " + FOLDER_NAME + "\n" +
    "폴더 URL: " + folder.getUrl()
  );
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Panel de Gestión')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
