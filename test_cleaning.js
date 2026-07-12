const fs = require('fs');
const https = require('https');

// The cleanHtml function replica from Google Apps Script
function cleanHtml(html) {
  if (!html) return '';
  let text = html;
  
  // Remove scripts, styles, head, SVGs and comments
  text = text.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  text = text.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
  text = text.replace(/<head\b[^<]*(?:(?!<\/head>)<[^<]*)*<\/head>/gi, '');
  text = text.replace(/<svg\b[^<]*(?:(?!<\/svg>)<[^<]*)*<\/svg>/gi, '');
  text = text.replace(/<!--[\s\S]*?-->/g, '');
  
  // Replace HTML tags with spaces
  text = text.replace(/<\/?[a-z][a-z0-9]*[^>]*>/gi, ' ');
  
  // Replace HTML entities
  text = text.replace(/&nbsp;/g, ' ')
             .replace(/&lt;/g, '<')
             .replace(/&gt;/g, '>')
             .replace(/&amp;/g, '&')
             .replace(/&quot;/g, '"');
             
  // Collapse whitespace
  text = text.replace(/\s+/g, ' ').trim();
  
  return text;
}

// Target test URL from the sheet: Pousada Recanto Caiçara
const testUrl = 'https://recantocaicara.com.br/';

console.log(`Fetching HTML from: ${testUrl}...`);

https.get(testUrl, (res) => {
  let data = '';
  res.on('data', (chunk) => {
    data += chunk;
  });
  
  res.on('end', () => {
    console.log(`\nOriginal HTML size: ${(data.length / 1024).toFixed(2)} KB`);
    
    const cleaned = cleanHtml(data);
    console.log(`Cleaned Text size: ${(cleaned.length / 1024).toFixed(2)} KB`);
    console.log(`Compression ratio: ${((1 - cleaned.length / data.length) * 100).toFixed(1)}% reduction!`);
    
    console.log('\n--- First 400 characters of cleaned text ---');
    console.log(cleaned.substring(0, 400));
    console.log('--------------------------------------------');
    
    console.log('\n--- Simulation of Gemini API prompt payload ---');
    const prompt = `Você é um monitor de preços automatizado.
Analise as informações do website do estabelecimento "Pousada Recanto Caiçara" (Atividade: HOSPEDAGEM).
Seu objetivo é encontrar preços de diária padrão, diária inicial, diária mínima ou tarifa "a partir de" para quarto duplo / casal.

Texto extraído do site:
"""
${cleaned.substring(0, 1000)}... [restante do site]
"""
`;
    console.log(prompt);
  });
}).on('error', (err) => {
  console.error("Error fetching site:", err.message);
});
