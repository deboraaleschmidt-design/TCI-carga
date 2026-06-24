const fs = require("fs");
const path = require("path");
const dir = __dirname;
const gs = fs.readFileSync(path.join(dir, "TCI-MontarPainelCarga.gs"), "utf8");
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const html = `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>Setup — Painel Carga TCI</title>
<style>
body{font-family:Segoe UI,sans-serif;max-width:960px;margin:24px auto;padding:0 16px;background:#f5f5f5;color:#222}
h1{color:#1565c0} .step{background:#fff;border-radius:8px;padding:16px;margin:12px 0;box-shadow:0 1px 3px #0002}
button{background:#1565c0;color:#fff;border:0;padding:10px 18px;border-radius:6px;cursor:pointer;font-size:14px;margin:4px 4px 4px 0}
button:hover{background:#0d47a1} textarea{width:100%;height:320px;font-family:Consolas,monospace;font-size:11px}
.info{background:#e3f2fd;padding:12px;border-radius:6px;font-size:13px}
</style></head><body>
<h1>Painel Carga TCI v2 — setup</h1>
<div class="info">Tudo no <strong>Chrome</strong>. Planilha nova abre em outra aba (CRIAR-PAINEL-V2.bat).</div>
<div class="step"><b>1</b> Renomeie a planilha: <em>PAINEL CARGA TCI v2</em></div>
<div class="step"><b>2</b> Extensões → Apps Script → apague o padrão → cole o código abaixo → Salvar
<br><button type="button" onclick="copiar()">Copiar código completo</button>
<textarea id="code" readonly>${esc(gs)}</textarea></div>
<div class="step"><b>3</b> F5 na planilha → menu <strong>TCI Carga → Montar estrutura inicial</strong> → autorize</div>
<div class="step"><b>4</b> Após colar BASE MICELANEAS (~10h): <strong>Registrar snapshot</strong> → <strong>Atualizar painel miscelânia</strong></div>
<div class="step"><b>Fontes (cópias)</b><br>
MATERIAIS 1dxK-p-blCB8TBRAfCBRWhjXHNiYBpGExkFAQE9PQa4M<br>
EQUIPES 1K1SzwTSSCri57o0U1A58EWT6cA4s0-XFjC5f5PyYR-A</div>
<script>
function copiar(){const t=document.getElementById('code');t.select();document.execCommand('copy');alert('Código copiado! Cole no Apps Script.');}
</script></body></html>`;
fs.writeFileSync(path.join(dir, "SETUP-PAINEL-CARGA.html"), html, "utf8");
console.log("SETUP-PAINEL-CARGA.html gerado");
