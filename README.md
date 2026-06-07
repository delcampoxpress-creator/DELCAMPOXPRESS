<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DEL CAMPO XPRESS — Sistema de Pedidos</title>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,300;0,9..144,700;0,9..144,900;1,9..144,400&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600;9..40,700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root {
  --cream: #F5F0E8;
  --dark: #1A1A18;
  --green: #2D5016;
  --green-light: #4A7C28;
  --green-pale: #E8F0DC;
  --amber: #C8872A;
  --amber-light: #F5C76A;
  --white: #FDFCFA;
  --border: #D4C9B0;
  --muted: #6B6455;
  --red: #C0392B;
  --shadow: 0 2px 16px rgba(0,0,0,0.09);
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'DM Sans',sans-serif;background:var(--cream);color:var(--dark);min-height:100vh;font-size:13px}

/* HEADER */
header{background:var(--dark);padding:14px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:200;border-bottom:3px solid var(--amber)}
.logo{font-family:'Fraunces',serif;font-size:18px;font-weight:900;color:var(--white);display:flex;align-items:center;gap:8px}
.logo span{color:var(--amber-light)}
.logo small{font-size:10px;font-weight:400;color:#888;font-family:'DM Sans',sans-serif;letter-spacing:1px;text-transform:uppercase;margin-left:4px}
.hbadge{background:var(--green);color:#fff;font-size:10px;font-weight:700;padding:4px 10px;border-radius:20px;letter-spacing:.5px;text-transform:uppercase}

/* TABS */
.tabs{background:#111;display:flex;padding:0 24px;gap:2px;border-bottom:1px solid #2a2a2a;overflow-x:auto;scrollbar-width:none}
.tabs::-webkit-scrollbar{display:none}
.tab{padding:10px 15px;font-size:11px;font-weight:600;color:#777;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap;text-transform:uppercase;letter-spacing:.5px}
.tab:hover{color:#bbb}
.tab.active{color:var(--amber-light);border-bottom-color:var(--amber-light)}

/* PANELS */
.panel{display:none;padding:24px;max-width:1300px;margin:0 auto}
.panel.active{display:block}
.ptitle{font-family:'Fraunces',serif;font-size:24px;font-weight:900;margin-bottom:4px}
.psub{color:var(--muted);font-size:12px;margin-bottom:20px}

/* CARDS */
.card{background:var(--white);border:1px solid var(--border);border-radius:12px;padding:20px;margin-bottom:16px;box-shadow:var(--shadow)}
.ctitle{font-family:'Fraunces',serif;font-size:14px;font-weight:700;margin-bottom:14px;display:flex;align-items:center;gap:7px}

/* GRID UTILS */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
.gfull{grid-column:1/-1}
@media(max-width:700px){.g2,.g3{grid-template-columns:1fr}}

/* FORM */
.fg{display:flex;flex-direction:column;gap:5px}
label{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
input,select,textarea{padding:9px 12px;border:1.5px solid var(--border);border-radius:8px;font-family:'DM Sans',sans-serif;font-size:13px;color:var(--dark);background:var(--white);outline:none;transition:border-color .2s;width:100%}
input:focus,select:focus,textarea:focus{border-color:var(--green)}
select{appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%236B6455' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 10px center;padding-right:28px}
textarea{resize:vertical;min-height:70px}
.fdivider{border:none;border-top:1px solid var(--border);margin:6px 0;grid-column:1/-1}
.flabel{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--green);grid-column:1/-1;padding-top:4px}

/* BUTTONS */
.btn{padding:9px 18px;border-radius:8px;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:700;cursor:pointer;border:none;transition:all .2s;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn-primary{background:var(--green);color:#fff}
.btn-primary:hover{background:var(--green-light)}
.btn-amber{background:var(--amber);color:#fff}
.btn-amber:hover{background:#B87820}
.btn-dark{background:var(--dark);color:#fff}
.btn-dark:hover{background:#333}
.btn-outline{background:transparent;color:var(--dark);border:1.5px solid var(--border)}
.btn-outline:hover{background:var(--cream)}
.btn-red{background:var(--red);color:#fff}
.btn-wa{background:#25D366;color:#fff}
.btn-wa:hover{background:#1fad55}
.btn-sm{padding:6px 11px;font-size:11px;border-radius:7px}
.brow{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}

/* ALERTS */
.alert{padding:10px 14px;border-radius:8px;font-size:12px;margin-bottom:10px;display:none;font-weight:600}
.alert.show{display:block}
.alert-ok{background:#D4EDDA;color:#155724;border:1px solid #C3E6CB}
.alert-err{background:#F8D7DA;color:#721C24;border:1px solid #F5C6CB}

/* PRODUCTS CATALOG */
.prod-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px}
.prod-card{background:var(--white);border:1.5px solid var(--border);border-radius:12px;padding:14px 16px;display:flex;align-items:center;gap:12px;transition:box-shadow .2s}
.prod-card:hover{box-shadow:0 4px 16px rgba(0,0,0,.1);border-color:var(--amber)}
.prod-emoji{font-size:28px;width:44px;height:44px;display:flex;align-items:center;justify-content:center;background:var(--cream);border-radius:10px;flex-shrink:0}
.prod-info{flex:1;min-width:0}
.prod-name{font-weight:700;font-size:13px;margin-bottom:2px}
.prod-unit{font-size:10px;color:var(--muted);margin-bottom:6px}
.prod-controls{display:flex;gap:5px;align-items:center;flex-wrap:wrap}
.price-field{width:90px;padding:5px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:12px;font-weight:600;color:var(--dark)}
.price-field:focus{border-color:var(--amber)}
.unit-sel{width:75px;padding:5px 6px;border:1.5px solid var(--border);border-radius:6px;font-size:11px;padding-right:20px}
.add-btn{padding:5px 10px;background:var(--green);color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;transition:background .15s}
.add-btn:hover{background:var(--green-light)}
.add-prod-new{border:2px dashed var(--green);background:var(--green-pale);border-radius:12px;padding:14px 16px;display:flex;align-items:center;gap:8px;cursor:pointer;transition:background .15s;font-size:12px;font-weight:600;color:var(--green)}
.add-prod-new:hover{background:#d2e8b4}

/* PROMOS */
.promo-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px;margin-bottom:18px}
.promo-card{background:var(--white);border:2px solid var(--border);border-radius:14px;overflow:hidden;transition:transform .2s,box-shadow .2s;cursor:default}
.promo-card:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,0,0,.1)}
.ph{padding:14px 16px;font-family:'Fraunces',serif;font-weight:900;font-size:16px;color:#fff;display:flex;justify-content:space-between;align-items:center}
.pb{padding:14px 16px}
.pitems{list-style:none;font-size:12px;margin-bottom:10px}
.pitems li{padding:2px 0;display:flex;gap:6px}
.pitems li::before{content:"•";color:var(--green);font-weight:700}
.pprice{font-family:'Fraunces',serif;font-size:22px;font-weight:900;color:var(--green)}
.pbtn{width:100%;margin-top:8px;padding:8px;background:var(--green);color:#fff;border:none;border-radius:8px;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:700;cursor:pointer;transition:background .2s}
.pbtn:hover{background:var(--green-light)}
.ph-junio{background:linear-gradient(135deg,var(--green),var(--green-light))}
.ph-bodegon{background:linear-gradient(135deg,#5D4037,#8D6E63)}
.ph-big{background:linear-gradient(135deg,#1565C0,#1976D2)}
.ph-gymbro{background:linear-gradient(135deg,#2E7D32,#43A047)}
.ph-campo{background:linear-gradient(135deg,var(--amber),#E67E22)}
.ph-cajon{background:linear-gradient(135deg,#37474F,#546E7A)}
.promo-price-row{display:flex;align-items:center;gap:8px}
.promo-price-input{width:110px;padding:5px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:13px;font-weight:700}
.new-promo-btn{border:2px dashed var(--amber);background:#fff8ee;border-radius:14px;padding:14px 16px;display:flex;align-items:center;justify-content:center;gap:8px;cursor:pointer;font-size:13px;font-weight:700;color:var(--amber);transition:background .15s}
.new-promo-btn:hover{background:#fef0cc}

/* SORTEO */
.sorteo-box{background:linear-gradient(135deg,var(--dark),#2D2D2A);border-radius:14px;padding:20px;text-align:center;color:#fff;border:2px solid var(--amber);margin-bottom:14px}
.sorteo-num{font-family:'Fraunces',serif;font-size:52px;font-weight:900;color:var(--amber-light);line-height:1;letter-spacing:-2px}
.sorteo-input{width:110px;padding:9px 12px;border:2px solid var(--amber);border-radius:8px;background:#2D2D2A;color:#fff;font-family:'Fraunces',serif;font-size:17px;font-weight:700;text-align:center;outline:none;margin-right:4px}

/* PAYMENT BOX */
.paybox{background:linear-gradient(135deg,#1A1A18,#2D2D2A);border-radius:10px;padding:16px 18px;color:#fff;display:flex;flex-direction:column;gap:5px}
.paybox .pt{font-size:10px;color:var(--amber-light);text-transform:uppercase;letter-spacing:1px;font-weight:700}
.paybox .pc{font-family:'Fraunces',serif;font-size:16px;font-weight:700;letter-spacing:1.5px}
.paybox .pa{font-size:12px;color:#aaa}
.paybox .pn{font-size:11px;color:var(--amber-light);margin-top:4px}

/* TICKET */
.ticket-preview{background:#fff;border:2px dashed #ccc;border-radius:12px;padding:24px;font-family:monospace;font-size:12px;line-height:1.7;white-space:pre-wrap;max-height:400px;overflow-y:auto;margin-bottom:12px}
.ticket-header{text-align:center;font-family:'Fraunces',serif;font-size:16px;font-weight:900;margin-bottom:12px;border-bottom:2px solid var(--dark);padding-bottom:10px}

/* TABLE */
.twrap{overflow-x:auto;margin-top:14px}
table{width:100%;border-collapse:collapse;font-size:11px;min-width:900px}
th{background:var(--dark);color:var(--amber-light);padding:9px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}
td{padding:9px 10px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:hover td{background:var(--green-pale)}
.sbadge{display:inline-block;padding:2px 7px;border-radius:20px;font-size:9px;font-weight:700;text-transform:uppercase}
.s-pendiente{background:#FFF3CD;color:#856404}
.s-confirmado{background:#D1ECF1;color:#0C5460}
.s-entregado{background:#D4EDDA;color:#155724}
.s-cancelado{background:#F8D7DA;color:#721C24}
.delbtn{color:var(--red);cursor:pointer;font-size:11px;border:none;background:none;padding:2px 5px}
.delbtn:hover{background:#fee;border-radius:4px}

/* STATS */
.stat-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px;margin-bottom:20px}
.scard{background:var(--white);border:1px solid var(--border);border-radius:12px;padding:18px;text-align:center;position:relative;overflow:hidden}
.scard::before{content:'';position:absolute;top:0;left:0;right:0;height:4px}
.scard.green::before{background:var(--green)}
.scard.amber::before{background:var(--amber)}
.scard.blue::before{background:#1565C0}
.scard.red::before{background:var(--red)}
.scard.purple::before{background:#7B1FA2}
.speriod{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:5px}
.sval{font-family:'Fraunces',serif;font-size:28px;font-weight:900;color:var(--dark);line-height:1}
.slabel{font-size:11px;color:var(--muted);margin-top:3px}

/* RANKING */
.rank-item{display:flex;align-items:center;justify-content:space-between;padding:10px 14px;background:var(--cream);border-radius:9px;margin-bottom:7px}
.rank-pos{width:24px;height:24px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;color:#fff;flex-shrink:0}
.rank-name{flex:1;font-weight:600;font-size:13px;margin:0 10px}
.rank-bar-bg{flex:2;height:8px;background:var(--border);border-radius:4px;overflow:hidden}
.rank-bar{height:100%;background:var(--green);border-radius:4px;transition:width .5s}
.rank-count{font-family:'Fraunces',serif;font-size:15px;font-weight:900;color:var(--dark);margin-left:10px}

/* MODAL */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:500;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--white);border-radius:16px;padding:28px;max-width:480px;width:90%;max-height:90vh;overflow-y:auto;box-shadow:0 16px 48px rgba(0,0,0,.25)}
.modal h2{font-family:'Fraunces',serif;font-size:20px;font-weight:900;margin-bottom:16px}
.modal-close{float:right;background:none;border:none;font-size:20px;cursor:pointer;color:var(--muted)}

/* EMPTY */
.empty{text-align:center;padding:40px 20px;color:var(--muted)}
.empty .ei{font-size:44px;margin-bottom:10px}

/* HIGHLIGHT */
.highlight-box{background:linear-gradient(135deg,var(--green-pale),#d8f0b8);border:1.5px solid var(--green);border-radius:12px;padding:16px;margin-bottom:14px}

/* ORDER SUMMARY CHIP */
.chip{display:inline-flex;align-items:center;gap:5px;background:var(--green-pale);color:var(--green);font-size:11px;font-weight:600;padding:3px 9px;border-radius:20px;margin:2px}
.chip-rm{background:none;border:none;cursor:pointer;color:var(--green);font-size:13px;padding:0;line-height:1}

/* TOTAL BAR */
.total-bar{background:var(--dark);color:#fff;border-radius:12px;padding:16px 20px;display:flex;justify-content:space-between;align-items:center;margin-top:14px}
.total-amount{font-family:'Fraunces',serif;font-size:28px;font-weight:900;color:var(--amber-light)}

/* SEARCH */
.search-bar{padding:9px 14px;border:1.5px solid var(--border);border-radius:8px;font-size:13px;width:100%;max-width:340px;margin-bottom:14px}
</style>
</head>
<body>

<header>
  <div class="logo">🐔 DEL CAMPO <span>XPRESS</span><small>Alimentos</small></div>
  <div class="hbadge">⚡ Sistema de Pedidos</div>
</header>

<div class="tabs">
  <div class="tab active" onclick="goTab('promos',this)">🏷️ Promos</div>
  <div class="tab" onclick="goTab('catalogo',this)">🛒 Catálogo</div>
  <div class="tab" onclick="goTab('pedido',this)">📋 Nuevo Pedido</div>
  <div class="tab" onclick="goTab('ticket',this)">🎫 Ticket</div>
  <div class="tab" onclick="goTab('planilla',this)">📊 Planilla</div>
  <div class="tab" onclick="goTab('stats',this)">📈 Estadísticas</div>
</div>

<div id="tab-promos" class="panel active">
  <div class="ptitle">🏷️ Promos Activas</div>
  <div class="psub">Hacé clic en "Agregar al pedido" para sumar una promo</div>
  <div class="promo-grid" id="promosGrid"></div>
  <div class="new-promo-btn" onclick="openNewPromoModal()">➕ Agregar nueva promo</div>
</div>

<div id="tab-catalogo" class="panel">
  <div class="ptitle">🛒 Catálogo de Productos</div>
  <div class="psub">Ajustá el precio y la cantidad antes de agregar al pedido</div>

  <div class="highlight-box">
    <b>🛍️ Pedido armado:</b>
    <div id="orderChips" style="margin-top:8px;min-height:28px"><span style="color:var(--muted);font-style:italic;font-size:12px">Ningún producto agregado aún…</span></div>
    <div class="total-bar" style="margin-top:12px">
      <div>
        <div style="font-size:12px;color:#aaa">Total del pedido</div>
        <div style="font-size:11px;color:#888;margin-top:1px" id="orderQtyLabel">0 items</div>
      </div>
      <div class="total-amount" id="orderTotal">$0</div>
    </div>
    <div class="brow" style="margin-top:10px">
      <button class="btn btn-primary" onclick="usarEnPedido()">📋 Usar en Pedido</button>
      <button class="btn btn-outline" onclick="clearOrder()">🔄 Limpiar</button>
    </div>
  </div>

  <div class="prod-grid" id="prodGrid"></div>
  <div style="margin-top:12px">
    <div class="add-prod-new" onclick="openNewProdModal()">➕ Agregar nuevo producto</div>
  </div>
</div>

<div id="tab-pedido" class="panel">
  <div class="ptitle">📋 Registrar Pedido</div>
  <div class="psub">Completá todos los datos del cliente</div>
  <div class="alert alert-ok" id="alertOk">✅ ¡Pedido registrado correctamente!</div>
  <div class="alert alert-err" id="alertErr">⚠️ Por favor completá los campos obligatorios.</div>

  <div style="display:grid;grid-template-columns:1fr 320px;gap:18px;align-items:start">
    <div>
      <div class="card">
        <div class="ctitle">👤 Datos del Cliente</div>
        <div class="g2">
          <div class="fg"><label>Nombre Completo *</label><input id="f-nombre" type="text" placeholder="Nombre y Apellido"></div>
          <div class="fg"><label>Teléfono / WhatsApp *</label><input id="f-tel" type="tel" placeholder="11 2345-6789"></div>
          <div class="fg gfull"><label>Dirección Completa *</label><input id="f-dir" type="text" placeholder="Calle, número, piso, barrio, localidad"></div>
          <hr class="fdivider">
          <div class="flabel">🚚 Datos de Entrega</div>
          <div class="fg"><label>Día de Entrega *</label>
            <select id="f-dia">
              <option value="">Seleccioná</option>
              <option>Miércoles</option>
              <option>Sábado</option>
            </select>
          </div>
          <div class="fg"><label>Rango Horario *</label>
            <select id="f-hora">
              <option value="">Seleccioná</option>
              <option>16:00 a 19:00</option>
              <option>17:00 a 20:00</option>
              <option>18:00 a 22:00</option>
            </select>
          </div>
          <hr class="fdivider">
          <div class="flabel">💳 Pago y Pedido</div>
          <div class="fg"><label>Forma de Pago *</label>
            <select id="f-pago">
              <option value="">Seleccioná</option>
              <option>Transferencia</option>
              <option>Efectivo</option>
            </select>
          </div>
          <div class="fg"><label>Monto Total ($)</label><input id="f-monto" type="number" placeholder="0"></div>
          <div class="fg gfull"><label>Productos / Promos</label><textarea id="f-productos" rows="3" placeholder="Promo Junio x1, Albóndigas 1kg x2…"></textarea></div>
          <div class="fg"><label>Estado</label>
            <select id="f-estado">
              <option>Pendiente</option>
              <option>Confirmado</option>
              <option>Entregado</option>
              <option>Cancelado</option>
            </select>
          </div>
          <div class="fg"><label>Notas</label><input id="f-notas" type="text" placeholder="Observaciones internas…"></div>
        </div>
        <div class="brow">
          <button class="btn btn-primary" onclick="registrarPedido()">✅ Registrar Pedido</button>
          <button class="btn btn-outline" onclick="limpiarForm()">🔄 Limpiar</button>
        </div>
      </div>
    </div>

    <div>
      <div class="sorteo-box">
        <div style="font-size:10px;text-transform:uppercase;letter-spacing:1px;color:var(--amber-light);font-weight:700;margin-bottom:10px">🎰 Número de Sorteo</div>
        <div class="sorteo-num" id="sorteoNum">—</div>
        <div style="display:flex;gap:6px;justify-content:center;margin-top:12px">
          <input type="number" class="sorteo-input" id="sorteoManual" placeholder="N°" min="1" oninput="document.getElementById('sorteoNum').textContent=this.value||'—'">
          <button class="btn btn-amber btn-sm" onclick="autoSorteo()">🎲 Auto</button>
        </div>
        <div style="font-size:10px;color:#888;margin-top:8px">Sorteo de fin de mes 🎁</div>
      </div>

      <div class="card">
        <div class="ctitle">🖥️ Datos de Pago</div>
        <div class="paybox">
          <div class="pt">💳 TRANSFERENCIA / MP</div>
          <div class="pc">0000003100083792527379</div>
          <div class="pa">ALIAS: campo.express.mp</div>
          <div class="pn">📸 Enviar comprobante al finalizar</div>
        </div>
      </div>

      <div class="card">
        <div class="ctitle">🎫 Ticket para el Cliente</div>
        <button class="btn btn-dark" style="width:100%" onclick="generarTicket()">🖨️ Generar Ticket</button>
      </div>
    </div>
  </div>
</div>

<div id="tab-ticket" class="panel">
  <div class="ptitle">🎫 Ticket del Cliente</div>
  <div class="psub">Generá el comprobante descargable para enviar al cliente</div>

  <div class="card">
    <div class="ctitle">📄 Vista previa del Ticket</div>
    <div class="ticket-preview" id="ticketPreview">
      <div style="text-align:center;color:#aaa;padding:20px">Complete el formulario del pedido y generá el ticket</div>
    </div>
    <div class="brow">
      <button class="btn btn-primary" onclick="descargarTicket()">⬇️ Descargar Ticket (.txt)</button>
      <button class="btn btn-wa" onclick="copiarTicket()">📋 Copiar para WhatsApp</button>
      <button class="btn btn-outline" onclick="generarTicket()">🔄 Actualizar</button>
    </div>
  </div>
</div>

<div id="tab-planilla" class="panel">
  <div class="ptitle">📊 Planilla de Pedidos</div>
  <div class="psub">Registro completo de todos los clientes y pedidos</div>

  <div class="brow" style="margin-bottom:14px">
    <button class="btn btn-primary btn-sm" onclick="exportExcel()">📥 Exportar Excel</button>
    <button class="btn btn-outline btn-sm" onclick="exportCSV()">📄 Exportar CSV</button>
    <button class="btn btn-outline btn-sm" onclick="renderTable()">🔄 Actualizar</button>
    <button class="btn btn-red btn-sm" onclick="clearAll()">🗑️ Limpiar Todo</button>
  </div>
  <input class="search-bar" id="searchBar" placeholder="🔍 Buscar cliente, producto, estado…" oninput="renderTable()">

  <div class="card" style="padding:0;overflow:hidden">
    <div class="twrap">
      <table>
        <thead><tr>
          <th>#</th><th>🎰 Sorteo</th><th>Nombre</th><th>Teléfono</th>
          <th>Dirección</th><th>Día</th><th>Horario</th>
          <th>Productos</th><th>Monto</th><th>Pago</th><th>Estado</th><th>Fecha</th><th></th>
        </tr></thead>
        <tbody id="pedidosTbody">
          <tr><td colspan="13"><div class="empty"><div class="ei">📋</div><p>No hay pedidos aún</p></div></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</div>

<div id="tab-stats" class="panel">
  <div class="ptitle">📈 Estadísticas</div>
  <div class="psub">Panel de control de ventas y pedidos</div>

  <div class="stat-grid">
    <div class="scard green"><div class="speriod">Esta semana</div><div class="sval" id="st-cl-sem">0</div><div class="slabel">👥 Clientes</div></div>
    <div class="scard amber"><div class="speriod">Esta semana</div><div class="sval" id="st-vt-sem">$0</div><div class="slabel">💰 Ventas</div></div>
    <div class="scard blue"><div class="speriod">Este mes</div><div class="sval" id="st-cl-mes">0</div><div class="slabel">👥 Clientes</div></div>
    <div class="scard green"><div class="speriod">Este mes</div><div class="sval" id="st-vt-mes">$0</div><div class="slabel">💰 Ventas</div></div>
    <div class="scard amber"><div class="speriod">Total acumulado</div><div class="sval" id="st-cl-tot">0</div><div class="slabel">👥 Clientes</div></div>
    <div class="scard purple"><div class="speriod">Total acumulado</div><div class="sval" id="st-vt-tot">$0</div><div class="slabel">💰 Ventas</div></div>
  </div>

  <div class="g2">
    <div class="card">
      <div class="ctitle">🏆 Promos Más Pedidas</div>
      <div id="rankPromos"><div class="empty"><div class="ei">🏷️</div><p>Sin datos aún</p></div></div>
    </div>
    <div class="card">
      <div class="ctitle">🥇 Productos Más Pedidos</div>
      <div id="rankProds"><div class="empty"><div class="ei">🛒</div><p>Sin datos aún</p></div></div>
    </div>
  </div>

  <div class="card">
    <div class="ctitle">📊 Pedidos por Estado</div>
    <div id="estadoStats" style="display:flex;gap:10px;flex-wrap:wrap"></div>
  </div>
</div>

<div class="modal-bg" id="modalProd">
  <div class="modal">
    <button class="modal-close" onclick="closeModal('modalProd')">✕</button>
    <h2>➕ Nuevo Producto</h2>
    <div style="display:flex;flex-direction:column;gap:12px">
      <div class="fg"><label>Nombre del producto *</label><input id="mp-nombre" type="text" placeholder="Ej: Suprema rellena"></div>
      <div class="fg"><label>Emoji (ícono)</label><input id="mp-emoji" type="text" placeholder="🍗" maxlength="4"></div>
      <div class="fg"><label>Precio de referencia ($)</label><input id="mp-precio" type="number" placeholder="0"></div>
      <div class="fg"><label>Unidad de venta</label>
        <select id="mp-unidad">
          <option value="kg">kg</option>
          <option value="unidad">Unidad</option>
          <option value="maple">Maple</option>
          <option value="docena">Docena</option>
        </select>
      </div>
    </div>
    <div class="brow" style="margin-top:16px">
      <button class="btn btn-primary" onclick="guardarNuevoProd()">💾 Guardar</button>
      <button class="btn btn-outline" onclick="closeModal('modalProd')">Cancelar</button>
    </div>
  </div>
</div>

<div class="modal-bg" id="modalPromo">
  <div class="modal">
    <button class="modal-close" onclick="closeModal('modalPromo')">✕</button>
    <h2>🏷️ Nueva Promo</h2>
    <div style="display:flex;flex-direction:column;gap:12px">
      <div class="fg"><label>Nombre de la promo *</label><input id="mpr-nombre" type="text" placeholder="Ej: Promo Familia"></div>
      <div class="fg"><label>Descripción / Contenido</label><textarea id="mpr-desc" rows="4" placeholder="2kg Pechuga&#10;1kg Milanesa Premium&#10;1 Maple de huevos"></textarea></div>
      <div class="fg"><label>Precio ($)</label><input id="mpr-precio" type="number" placeholder="0"></div>
      <div class="fg"><label>Color del encabezado</label>
        <select id="mpr-color">
          <option value="ph-junio">Verde</option>
          <option value="ph-bodegon">Marrón</option>
          <option value="ph-big">Azul</option>
          <option value="ph-gymbro">Verde oscuro</option>
          <option value="ph-campo">Ámbar</option>
          <option value="ph-cajon">Gris</option>
        </select>
      </div>
    </div>
    <div class="brow" style="margin-top:16px">
      <button class="btn btn-primary" onclick="guardarNuevaPromo()">💾 Guardar Promo</button>
      <button class="btn btn-outline" onclick="closeModal('modalPromo')">Cancelar</button>
    </div>
  </div>
</div>

<script>
// El código JavaScript original que maneja la lógica de la app va aquí...
</script>
</body>
</html>
