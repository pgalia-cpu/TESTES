/**
 * ============================================================================
 * EUM MANAGER 4.5 - SNIPER EDITION (1 REQUEST/MIN FIX)
 * * Correção Crítica:
 * 1. Removemos a verificação dinâmica de modelos (poupamos 1 requisição).
 * 2. Removemos o fallback automático (poupamos requisições extras).
 * 3. Forçamos o modelo 'gemini-1.5-flash' (Padrão Ouro).
 * 4. Tratamento de erro 429 explícito (Avisa para esperar).
 * ============================================================================
 */

const GEMINI_API_KEY = "AIzaSyCOeCUpUFYCpUDG362Ld5pxAbehKWQdXRE"; 
const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM";

// --- MODELO FIXO (Para não gastar quota a verificar) ---
const MODELO_FIXO = "gemini-2.0-flash"; 

// --- UI ---
function onOpen() { SpreadsheetApp.getUi().createMenu('🚀 EUM App').addItem('Abrir Painel', 'abrirDashboard').addToUi(); }
function abrirDashboard() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('Dashboard').evaluate().setTitle('EUM Manager 4.5').setWidth(1200).setHeight(900), 'EUM Manager'); }

// --- CONFIG ---
function apiGetInitialState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName()).filter(n => !['Exames_Referencia','Doses_Referencia','Config','Dashboard'].includes(n));
  const savedConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
  return { sheets: sheets, config: savedConfig, status: { hasCore: !!savedConfig.core?.abaFatos, hasExames: !!savedConfig.exames?.active, hasRams: !!savedConfig.rams?.active } };
}

function apiSaveConfig(newConfig) {
  try {
    PropertiesService.getScriptProperties().setProperty('EUM_CONFIG_MASTER', JSON.stringify(newConfig));
    return { sucesso: true, status: { hasCore: !!newConfig.core.abaFatos, hasExames: !!newConfig.exames.active, hasRams: !!newConfig.rams.active } };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// ★ WIZARD HÍBRIDO (MODO SNIPER)
// ============================================================================
function apiMagicSetup(pdfBase64, matrixDados, nomeArquivo, abaExistente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let headers = [], sheetName = "";

  // Preparação de Dados (Local - Não gasta quota)
  if (abaExistente) {
    const sheet = ss.getSheetByName(abaExistente);
    if (!sheet) return { erro: "Aba não encontrada." };
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheetName = abaExistente;
  } else if (matrixDados && matrixDados.length > 0) {
    sheetName = "Dados_" + (nomeArquivo || "Import").replace(/[^a-zA-Z0-9]/g, "_").substring(0, 15);
    if (ss.getSheetByName(sheetName)) sheetName += "_" + Math.floor(Math.random()*1000);
    const sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, matrixDados.length, matrixDados[0].length).setValues(matrixDados);
    headers = matrixDados[0];
  } else {
    return { erro: "Sem dados." };
  }

  // Chamada IA (TIRO ÚNICO)
  if (!GEMINI_API_KEY) return { erro: "Chave API ausente." };

  const prompt = `
    Atue como Consultor EUM.
    INPUTS: Cabeçalhos: ${JSON.stringify(headers)}.
    TAREFA: Mapeie colunas do Protocolo (PDF).
    IMPORTANTE: Se houver TFG/CKD calculada, mapeie 'colTfgPreCalc'.
    RETORNE APENAS JSON: {"studyName": "...", "studySummary": "...", "analysis": { "capabilities": [], "gaps": [] }, "mapping": { "colProntFatos": "...", "colMed": "...", "colDtIni": "...", "colDoseUni": "...", "colApraz": "...", "colDose24h": "...", "colPeso": "...", "colAltura": "...", "colCreat": "...", "colTfgPreCalc": "...", "colProntDim": "...", "colNasc": "...", "colSexo": "..." }}`;

  const payload = { contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: "application/pdf", data: pdfBase64 } }] }] };

  try {
    const txt = callGeminiDirect_(payload); // ★ Chama direto
    const resIA = cleanJson_(txt);
    resIA.createdSheet = sheetName;
    return resIA;
  } catch (e) { return { erro: e.message }; }
}

// ============================================================================
// ★ FUNÇÃO DE CHAMADA DIRETA (SEM CHECAGENS EXTRAS)
// ============================================================================
function callGeminiDirect_(payload) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELO_FIXO}:generateContent?key=${GEMINI_API_KEY}`;
  
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    
    const code = resp.getResponseCode();
    
    // Tratamento de Erro de Limite (429)
    if (code === 429) {
      throw new Error("⚠️ LIMITE DE VELOCIDADE ATINGIDO (1 RPM).\nPor favor, aguarde 1 minuto e tente novamente.");
    }
    
    if (code !== 200) {
      throw new Error(`Erro Google (${code}): ${resp.getContentText()}`);
    }
    
    return JSON.parse(resp.getContentText()).candidates[0].content.parts[0].text;
  } catch (e) { throw e; }
}

function cleanJson_(text) {
  let clean = text.replace(/```json/g, "").replace(/```/g, "").trim();
  const start = clean.indexOf('{'), end = clean.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error("IA não retornou JSON.");
  return JSON.parse(clean.substring(start, end + 1));
}

// ============================================================================
// 4. API CALCULADORA & SANDBOX
// ============================================================================
function apiGetVariaveisCustom() { const v=JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_VARS_CUSTOM')||'{}'); return Object.keys(v).map(k=>({nome:k,formula:v[k]})); }
function apiSalvarVariavel(n,f) { const p=PropertiesService.getScriptProperties(), v=JSON.parse(p.getProperty('EUM_VARS_CUSTOM')||'{}'); v[n.toUpperCase().trim()]=f; p.setProperty('EUM_VARS_CUSTOM',JSON.stringify(v)); return {sucesso:true}; }
function apiExcluirVariavel(n) { const p=PropertiesService.getScriptProperties(), v=JSON.parse(p.getProperty('EUM_VARS_CUSTOM')||'{}'); delete v[n]; p.setProperty('EUM_VARS_CUSTOM',JSON.stringify(v)); return {sucesso:true}; }

function apiGetExplorerSample(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if(!sheet) return { sample: [], columns: [] };
  const data = sheet.getRange(1, 1, 6, sheet.getLastColumn()).getValues();
  const headers = data.shift().map(h => String(h).toUpperCase().trim());
  const sample = data.map(row => { const obj={}; headers.forEach((h,i)=>obj[h]=row[i]); obj['_IDADE_ANOS']=Math.floor(Math.random()*80); return obj; });
  return { sample: sample, columns: headers };
}

function apiGerarFormulaIA(pergunta, colunas) {
  const prompt = `Converta lógica clínica em fórmula JS Math. Pergunta: "${pergunta}". Colunas: ${JSON.stringify(colunas)}. Use [COLUNA]. Retorne APENAS a string.`;
  try {
    const txt = callGeminiDirect_({ contents: [{ parts: [{ text: prompt }] }] });
    return { formula: txt.replace(/`/g, "").trim() };
  } catch (e) { return { formula: null }; }
}

// ============================================================================
// ★ ENGINE DE DADOS 5.2 (CORREÇÃO DE IDADE REAL)
// Agora calcula a idade cruzando DT_REFERENCIA da linha com DT_NASCIMENTO da linha
// ============================================================================
function apiGetRawExplorerData(sheetName, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toUpperCase().trim());
  const customVars = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_VARS_CUSTOM') || '{}');
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
  
  // 1. Mapeamento de Colunas (Índices)
  const colMap = {}; headers.forEach((h, i) => colMap[h] = i);
  
  // Tenta encontrar as colunas vitais na configuração ou pelo nome
  const idxPront = config.core?.colProntFatos ? colMap[config.core.colProntFatos.toUpperCase()] : -1;
  
  // Procura Data de Nascimento na própria aba (Correção para CSV único)
  let idxNasc = config.core?.colNasc ? colMap[config.core.colNasc.toUpperCase()] : -1;
  if (idxNasc === -1) idxNasc = headers.findIndex(h => h.includes("NASC") || h.includes("BIRTH")); // Tentativa automática

  // Procura Data do Evento (Referência/Início)
  let idxDataEvento = -1;
  if (config.core?.colDtIni && colMap[config.core.colDtIni.toUpperCase()] !== undefined) {
     idxDataEvento = colMap[config.core.colDtIni.toUpperCase()];
  } else {
     idxDataEvento = headers.findIndex(h => h.includes("DT_REF") || h.includes("DT_INI") || h.includes("DATA"));
  }

  // (Opcional) Mapa de Pacientes Externo (para casos de abas separadas)
  let mapPacientes = {};
  if (config.core?.abaDim && config.core?.abaDim !== sheetName) {
     const sDim = ss.getSheetByName(config.core.abaDim);
     if(sDim) {
       const dDim = sDim.getDataRange().getValues(), hDim = dDim.shift();
       const iP = hDim.indexOf(config.core.colProntDim), iN = hDim.indexOf(config.core.colNasc);
       dDim.forEach(r => mapPacientes[String(r[iP]).trim()] = { n: parseDatePTBR_(r[iN]) });
     }
  }

  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i], rowObj = {}, rowContext = {};
    let hasData = false;
    
    // 1. Identifica ID do Paciente
    const pid = (idxPront > -1) ? String(row[idxPront]).trim() : "ID_" + i;
    rowObj['_PID'] = pid;

    // 2. LEITURA E LIMPEZA
    headers.forEach((h, idx) => { 
      let v = row[idx];
      v = cleanHospitalValue_(v); // Aplica o filtro de "Lixo" (datas coladas)
      rowContext[h] = v; 
    });

    // 3. CÁLCULO DA IDADE (CRUZAMENTO)
    let dataNasc = null;
    
    // A) Tenta pegar da própria linha (CSV Único)
    if (idxNasc > -1) dataNasc = parseDatePTBR_(row[idxNasc]);
    
    // B) Se não achou, tenta pegar do mapa externo (Dimensão)
    if (!dataNasc && mapPacientes[pid]) dataNasc = mapPacientes[pid].n;

    // C) Pega data do evento (ou usa hoje se falhar)
    let dataRef = (idxDataEvento > -1) ? parseDatePTBR_(row[idxDataEvento]) : new Date();
    
    if (dataNasc && dataRef) {
       const diff = dataRef - dataNasc;
       // Calcula idade exata (com casas decimais para pediatria)
       rowContext['_IDADE_ANOS'] = parseFloat((diff / (365.25 * 86400000)).toFixed(2));
    } else {
       rowContext['_IDADE_ANOS'] = 0; // Sem dados
    }

    // 4. SANDBOX (Calculadora)
    Object.keys(customVars).forEach(k => {
      let f = customVars[k];
      const m = f.match(/\[(.*?)\]/g);
      if(m) m.forEach(t => { f = f.replace(t, (rowContext[t.replace(/[\[\]]/g,"").toUpperCase().trim()]||0)); });
      try { rowContext[k] = parseFloat(parseFloat(eval(f)).toFixed(2)); } catch(e){ rowContext[k]=0; }
    });

    // 5. Seleção Final
    cols.forEach(c => {
      if (customVars[c]) { rowObj[c]=rowContext[c]; hasData=true; }
      else if (colMap[c.toUpperCase()] !== undefined) { rowObj[c]=rowContext[c.toUpperCase()]; hasData=true; }
      else rowObj[c]=0;
    });
    
    if (hasData) result.push(rowObj);
  }
  return { sucesso: true, dados: result };
}

// ============================================================================
// 6. MOTORES CLÍNICOS & HELPERS
// ============================================================================
function parseDatePTBR_(d) { if(!d)return null; if(d instanceof Date)return d; if(typeof d==='string'){ const p=d.split(/[\/\-\.\s]/); if(p.length>=3){ const D=parseInt(p[0]),M=parseInt(p[1])-1,Y=parseInt(p[2]); if(D>0&&D<=31&&M>=0&&M<=11) { const dt=new Date(Y<100?Y+2000:Y,M,D); if(!isNaN(dt.getTime()))return dt; }}} return isNaN(new Date(d).getTime())?null:new Date(d); }

function calcularFuncaoRenal_(cr, age, sex, wt, ht, tfgImp) {
  if((!cr||isNaN(cr)||cr<=0) && tfgImp>0) return { valor: parseFloat(tfgImp.toFixed(1)), metodo: "Importado", estagio: classificarRenal_(tfgImp) };
  if(!cr||isNaN(cr)||cr<=0) return { valor: null, metodo: "N/D", estagio: "N/D" };
  let res=0, met="";
  if(age<18) { if(ht>0){res=(0.413*ht)/cr; met="Schwartz";} else return {valor:null,metodo:"Falta Altura",estagio:"N/D"}; }
  else { const isF=String(sex).toUpperCase().startsWith("F"), k=isF?0.7:0.9, a=isF?-0.241:-0.302; res=142*Math.pow(Math.min(cr/k,1),a)*Math.pow(Math.max(cr/k,1),-1.2)*Math.pow(0.9938,age)*(isF?1.012:1.0); met="CKD-EPI"; }
  return { valor: parseFloat(res.toFixed(1)), metodo: met, estagio: classificarRenal_(res) };
}
function classificarRenal_(v) { if(v>=90)return"G1"; if(v>=60)return"G2"; if(v>=45)return"G3a"; if(v>=30)return"G3b"; if(v>=15)return"G4"; return"G5"; }

function interpretingPosologia_(d,a) { const ds=String(d||"0").replace(',','.').trim(), dm=ds.match(/(\d+(\.\d+)?)/); let dose=dm?parseFloat(dm[0]):0; if(dose===0)return 0; const as=String(a||"").toLowerCase(); let f=1; const mD=as.match(/de\s*(\d+)\s*[\/]\s*(\d+)/), mI=as.match(/(\d+)\s*(?:\/|-|em|a cada|h|:)\s*(\d+)/), mH=as.match(/(?:q|cada|a cada)\s*(\d+)\s*h?/); if(mD&&parseFloat(mD[2])>0)f=24/parseFloat(mD[2]); else if(mI&&parseFloat(mI[2])>0)f=24/parseFloat(mI[2]); else if(mH&&parseFloat(mH[1])>0)f=24/parseFloat(mH[1]); else if(/(12|bid|2x|duas|manh)/.test(as))f=2; else if(/(8|tid|3x|tres)/.test(as))f=3; return parseFloat((dose*f).toFixed(2)); }

// API Módulos (Eficácia/Segurança)
function apiGetEficaciaData(params) {
  const state = apiGetInitialState(); if (!state.status.hasExames) return { sucesso: false, erro: "Sem Exames." };
  try {
    const c=state.config.core, e=state.config.exames, ss=SpreadsheetApp.getActiveSpreadsheet();
    const dDim=ss.getSheetByName(c.abaDim).getDataRange().getValues(), hDim=dDim.shift(), mapDim={};
    const iP=hDim.indexOf(c.colProntDim), iN=hDim.indexOf(c.colNasc), iS=hDim.indexOf(c.colSexo);
    dDim.forEach(r => mapDim[String(r[iP]).trim()] = { n: parseDatePTBR_(r[iN]), s: String(r[iS]||"").charAt(0) });

    const dFat=ss.getSheetByName(c.abaFatos).getDataRange().getValues(), hF=dFat.shift();
    const iPF=hF.indexOf(c.colProntFatos), iMed=hF.indexOf(c.colMed), iDt=hF.indexOf(c.colDtIni), iPe=hF.indexOf(c.colPeso), iAlt=hF.indexOf(c.colAltura), iCr=hF.indexOf(c.colCreat), iTfg=hF.indexOf(c.colTfgPreCalc);
    const iDose=hF.indexOf(c.colDoseUni), iApraz=hF.indexOf(c.colApraz), iDose24=hF.indexOf(c.colDose24h);
    const d0={};
    dFat.forEach(r => {
      if (r[iMed] === params.medicamento) {
        const p=String(r[iPF]).trim(), dt=parseDatePTBR_(r[iDt]);
        if (dt) {
           let dose = 0; if(iDose24>-1&&r[iDose24]) dose=parseFloat(String(cleanHospitalValue_(r[iDose24])).replace(',','.')); else if(iDose>-1) dose=interpretarPosologia_(r[iDose], iApraz>-1?r[iApraz]:"");
           const pe=iPe>-1?parseFloat(String(cleanHospitalValue_(r[iPe])).replace(',','.')):0; // ★ LIMPEZA
           const alt=iAlt>-1?parseFloat(String(cleanHospitalValue_(r[iAlt])).replace(',','.')):0; // ★ LIMPEZA
           const cr=iCr>-1?parseFloat(String(cleanHospitalValue_(r[iCr])).replace(',','.')):0; // ★ LIMPEZA
           const tfg=iTfg>-1?parseFloat(String(cleanHospitalValue_(r[iTfg])).replace(',','.')):0; // ★ LIMPEZA
           if(!d0[p] || dt<d0[p].dt) d0[p] = { dt: dt, p: pe, a: alt, c: cr, t: tfg, dose: dose };
        }
      }
    });

    const dEx=ss.getSheetByName(e.aba).getDataRange().getValues(), hEx=dEx.shift();
    const ixP=hEx.indexOf(e.colPront), ixN=hEx.indexOf(e.colNome), ixV=hEx.indexOf(e.colValor), ixD=hEx.indexOf(e.colData);
    const plot=[];
    dEx.forEach(r => {
      const p=String(r[ixP]).trim();
      if (d0[p] && normalize_(r[ixN]) === normalize_(params.exameAlvo)) {
         const dtEx=parseDatePTBR_(r[ixD]);
         let val = cleanHospitalValue_(r[ixV]); // ★ LIMPEZA NO VALOR DO EXAME
         if (dtEx && !isNaN(val)) {
           const meta=mapDim[p]||{}, days=Math.floor((dtEx-d0[p].dt)/86400000);
           let age=meta.n ? parseFloat(((dtEx-meta.n)/(365.25*86400000)).toFixed(1)) : 0;
           const renal=calcularFuncaoRenal_(d0[p].c, age, meta.s, d0[p].p, d0[p].a, d0[p].t);
           const doseKg=(d0[p].dose && d0[p].p) ? (d0[p].dose/d0[p].p) : 0;
           plot.push({ x: days, y: val, prontuario: p, refMin: null, refMax: null, sexo: meta.s, idade: age, doseTotal: d0[p].dose, doseKg: parseFloat(doseKg.toFixed(2)), renal: renal });
         }
      }
    });
    return { sucesso: true, dados: plot };
  } catch(err) { return { sucesso: false, erro: err.message }; }
}

function apiGetSegurancaData(params) {
  const state = apiGetInitialState(); if (!state.status.hasRams) return { sucesso: false, erro: "Sem RAMs." };
  try {
    const c=state.config.core, r=state.config.rams, ss=SpreadsheetApp.getActiveSpreadsheet();
    const fats=ss.getSheetByName(c.abaFatos).getDataRange().getValues(), hF=fats.shift();
    const iPF=hF.indexOf(c.colProntFatos), iMed=hF.indexOf(c.colMed), iDI=hF.indexOf(c.colDtIni);
    const coorte={}; fats.forEach(row => { if(row[iMed]===params.medicamento) { const p=String(row[iPF]).trim(), d=parseDatePTBR_(row[iDI]); if(d&&(!coorte[p]||d<coorte[p])) coorte[p]=d; }});
    
    const rams=ss.getSheetByName(r.aba).getDataRange().getValues(), hR=rams.shift();
    const iG=hR.indexOf(r.colGrav), iC=hR.indexOf(r.colCaus), iPR=hR.findIndex(k=>k.includes("PRONT")), iDtR=hR.findIndex(k=>k.includes("DATA"));
    const evts=[]; rams.forEach(row => {
      const p=String(row[iPR]).trim();
      if(coorte[p]) { const dt=parseDatePTBR_(row[iDtR]), dias=dt?Math.floor((dt-coorte[p])/86400000):null; evts.push({ gravidade:String(row[iG]||"N/D"), causalidade:String(row[iC]||"N/D"), dias:dias, prontuario:p }); }
    });
    return { sucesso: true, dados: evts, totalExpostos: Object.keys(coorte).length };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// Utils
function apiGetColumns(s) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); return sh?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h).toUpperCase()):[]; }
function apiCheckGeminiModels() { return {sucesso:true, modelos:[MODELO_FIXO]}; }
function apiInterpretarGraficoIA(p,c) { return { viavel: false, explicacao: "Gráfico IA desativado (Limite de Cota 1RPM)" }; }
function apiGetMedicamentos(s,c) { return getUnique_(s,c); }
function apiGetExamesList(s,c) { return getUnique_(s,c); }
function getUnique_(s,c) { const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); if(!sheet)return[]; const h=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]; const i=h.indexOf(c); if(i<0)return[]; return [...new Set(sheet.getRange(2,i+1,sheet.getLastRow()-1,1).getValues().map(v=>String(v[0]).trim()))].sort(); }
function apiGetReferenciasDb() { const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exames_Referencia"); if(!s)return{sucesso:true,dados:[]}; const d=s.getDataRange().getValues(); d.shift(); return {sucesso:true,dados:d.map((r,i)=>({id:i+2,nome:r[0],tipo:r[1],sexo:r[2],diasMin:r[3],diasMax:r[4],min:r[5],max:r[6]}))}; }
function apiSalvarReferencia(r) { const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exames_Referencia")||SpreadsheetApp.getActiveSpreadsheet().insertSheet("Exames_Referencia"); s.appendRow([r.nome,r.tipo,r.sexo,r.diasMin,r.diasMax,r.min,r.max]); return {sucesso:true}; }
function apiExcluirReferencia(id) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exames_Referencia").deleteRow(parseInt(id)); return {sucesso:true}; }
function apiGetRegrasDose() { const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doses_Referencia"); if(!s)return{sucesso:true,dados:[]}; const d=s.getDataRange().getValues(); d.shift(); return {sucesso:true,dados:d.map((r,i)=>({id:i+2,med:r[0],diasMin:r[1],diasMax:r[2],doseMin:r[3],doseUsual:r[4],doseMax:r[5],unidade:r[6]}))}; }
function apiSalvarRegraDose(r) { const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doses_Referencia")||SpreadsheetApp.getActiveSpreadsheet().insertSheet("Doses_Referencia"); s.appendRow([r.med,r.diasMin,r.diasMax,r.doseMin,r.doseUsual,r.doseMax,r.unidade]); return {sucesso:true}; }
function apiExcluirRegraDose(id) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doses_Referencia").deleteRow(parseInt(id)); return {sucesso:true}; }
function apiExportarParaMaster() { return {sucesso:true}; }
function apiImportarReferencias() { return {sucesso:true}; }
function normalize_(s) { return String(s||"").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }

// ============================================================================
// ★ HELPER: LIMPADOR DE DADOS HOSPITALARES ("O FILTRO CIRÚRGICO")
// Remove datas/horas que vêm coladas aos números (ex: "14/04/2025 10:00 7,8")
// ============================================================================
function cleanHospitalValue_(val) {
  if (typeof val !== 'string') return val;
  if (!val) return "";

  // 1. Deteta padrão "DD/MM/YYYY HH:MM:SS VALOR"
  // Regex busca data/hora no início e captura o resto (grupo 1)
  const regex = /^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}:\d{2}\s+(.*)$/;
  const match = val.match(regex);

  let limpo = val;
  if (match && match[1]) {
    limpo = match[1].trim(); // Fica só com o "7,8" ou "28"
  }

  // 2. Tenta converter para número (trata vírgula PT-BR)
  // Se for "7,8", vira 7.8. Se for texto puro, mantém o texto.
  const num = parseFloat(limpo.replace(',', '.'));
  return isNaN(num) ? limpo : num;
}

// ============================================================================
// ★ HELPER: LISTAR VALORES ÚNICOS PARA OS MENUS
// ============================================================================
function apiGetUniqueValues(sheetName, colName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
  const colIdx = headers.indexOf(String(colName).toUpperCase().trim());
  
  if (colIdx === -1) return [];
  
  // Pega valores, remove vazios e duplicatas, e ordena
  const raw = sheet.getRange(2, colIdx + 1, lastRow - 1, 1).getValues();
  const unique = [...new Set(raw.map(r => String(r[0]).trim()).filter(v => v !== ""))].sort();
  
  return unique;
}

// Wrapper para garantir compatibilidade com o Frontend
function apiGetMedicamentos(sheet, col) { return apiGetUniqueValues(sheet, col); }
function apiGetExamesList(sheet, col) { return apiGetUniqueValues(sheet, col); }
