// Lectura de la póliza y la factura de La Segunda a partir del texto de los PDF.
// Reconstruye las líneas por posición y busca los rótulos de cada dato.
export function linesFromText(tc){
  const items = tc.items.filter(it=>it.str && it.str.trim()).map(it=>({s:it.str, x:it.transform[4], y:it.transform[5], w:it.width}));
  items.sort((a,b)=> b.y-a.y || a.x-b.x);
  const lines=[];
  for(const it of items){
    let L = lines.find(l=>Math.abs(l.y-it.y)<2.5);
    if(!L){ L={y:it.y, items:[]}; lines.push(L); }
    L.items.push(it);
  }
  lines.sort((a,b)=>b.y-a.y);
  return lines.map(l=>{
    l.items.sort((a,b)=>a.x-b.x);
    let out="", end=null;
    for(const it of l.items){
      if(end!==null){ const gap=it.x-end; out += gap>12 ? "   " : (gap>1.5 ? " " : ""); }
      out += it.s; end = it.x+it.w;
    }
    return {y:l.y, text:out, items:l.items};
  });
}
const pnum = s => parseFloat(String(s).replace(/[$\s]/g,"").replace(/,/g,""));
function field(lines, label){
  for(const l of lines){
    const i = l.text.indexOf(label);
    if(i<0) continue;
    return l.text.slice(i+label.length).replace(/^\s+/,"").split(/\s{3,}/)[0].trim();
  }
  return null;
}
function certNumber(lines){
  for(const l of lines){
    const h = l.items.find(it=>/Certificado N/.test(it.s));
    if(!h) continue;
    const cx = h.x + h.w/2; let best=null;
    for(const l2 of lines){
      if(l2.y>=l.y || l.y-l2.y>40) continue;
      for(const it of l2.items){
        if(/^\d{1,5}$/.test(it.s.trim())){
          const d=Math.abs(it.x+it.w/2-cx);
          if(d<45 && (!best||d<best.d)) best={d, n:parseInt(it.s)};
        }
      }
    }
    return best? best.n : null;
  }
  return null;
}
export async function pageText(pdf,p){ const page=await pdf.getPage(p); const lines=linesFromText(await page.getTextContent()); page.cleanup&&page.cleanup(); return lines; }
export async function parsePoliza(pdf, onProgress){
  const out={vehiculos:[], meta:{}};
  for(let p=1;p<=pdf.numPages;p++){
    const lines = await pageText(pdf,p);
    const all = lines.map(l=>l.text).join("\n");
    if(p===1){
      const m=all.match(/(\d{1,3}\.\d{3}\.\d{3})/); if(m) out.meta.poliza=m[1];
      const fs=all.match(/(\d{2}\/\d{2}\/\d{4})[\s\S]*?(\d{2}\/\d{2}\/\d{4})/); if(fs){out.meta.desde=fs[1]; out.meta.hasta=fs[2];}
      const tc=all.match(/TOTAL CERTIFICADOS:\s*(\d+)/); if(tc) out.meta.totalCert=parseInt(tc[1]);
      const pf=all.match(/Premio final:\s*([\d,\.]+)/); if(pf) out.meta.premioFinal=pnum(pf[1]);
      const ti=lines.findIndex(l=>/Información del tomador/.test(l.text)); if(ti>=0&&lines[ti+1]) out.meta.tomador=lines[ti+1].text.split(/\s{3,}/)[0].trim();
    }
    if(all.includes("Valor del vehículo") && all.includes("Marca completa")){
      const val=field(lines,"Valor del vehículo"), marca=field(lines,"Marca completa"), pat=field(lines,"Patente"), anio=field(lines,"Año");
      const pm=all.match(/PREMIO TOTAL:\s*([\d,\.]+)/);
      out.vehiculos.push({cert:certNumber(lines), patente:(pat||"").replace(/[^A-Za-z0-9]/g,"").toUpperCase(), marca:marca||"", anio:anio||"", valor:val?pnum(val)||0:0, premio:pm?pnum(pm[1]):0});
    }
    if(onProgress && (p%8===0||p===pdf.numPages)){ onProgress(p, pdf.numPages, out.vehiculos.length); await new Promise(r=>setTimeout(r,0)); }
  }
  return out;
}
export function parseFacturaLines(lines){
  const all = lines.map(l=>l.text).join("\n");
  const g = re => { const m=all.match(re); return m? m[1] : null; };
  const amt = re => { const v=g(re); return v==null?null:pnum(v); };
  return {
    numero: g(/N°:\s*([\d-]+)/), fecha: g(/Fecha:\s*(\d{2}\/\d{2}\/\d{4})/),
    detalle: (g(/(AUTOMOTORES[^\n]*?)(\s{3,}|\n|$)/) || "").trim(),
    cae: g(/C\.A\.E\.:\s*(\d+)/),
    neto: amt(/BONIFICACI[ÓO]N:[\s\S]*?SUBTOTAL:\s*\$\s*([\d\.,]+)/) ?? amt(/SUBTOTAL:\s*\$\s*([\d\.,]+)/),
    iva: amt(/\nI\.V\.A\.:\s*\$\s*([\d\.,]+)/) ?? amt(/I\.V\.A\.:\s*\$\s*([\d\.,]+)/),
    perc: amt(/PERCEPCIONES:\s*\$\s*([\d\.,]+)/) ?? 0,
    otros: amt(/OTROS IMPUESTOS:\s*\$\s*([\d\.,]+)/) ?? 0,
    total: amt(/(?<![A-Z])TOTAL:\s*\$\s*([\d\.,]+)/)
  };
}

