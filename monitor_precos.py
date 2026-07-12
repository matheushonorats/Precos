import os
import re
import csv
import sys
import time
import requests
from bs4 import BeautifulSoup
import urllib3
from urllib.parse import urljoin, urlparse

# Configure UTF-8 encoding for console output on Windows
sys.stdout.reconfigure(encoding='utf-8')

# Disable SSL verification warnings for legacy/self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Request configurations
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}
TIMEOUT = 10
MAX_RETRIES = 2
DELAY_BETWEEN_SITES = 2.0  # seconds to avoid rate limits

CSV_PATH = 'estabelecimentos_para_monitorar.csv'

def clean_text(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    # Remove scripts, styles, SVGs, and noscript (keep header/footer/nav as they may contain promo prices)
    for element in soup(["script", "style", "svg", "noscript"]):
        element.decompose()
    text = soup.get_text(separator=' ')
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def is_internal(base_url, link_url):
    base_netloc = urlparse(base_url).netloc
    link_netloc = urlparse(link_url).netloc
    return not link_netloc or link_netloc == base_netloc

def extract_price_candidates(text):
    # Match standard Brazilian currency formats (e.g., R$ 150,00, R$150,00 or just 150,00)
    pattern = r'(?:R\$\s*)?(\b\d{1,3}(?:\.\d{3})*(?:,\d{2})\b)'
    matches = re.finditer(pattern, text)
    
    candidates = []
    seen = set()
    for match in matches:
        val_str = match.group(1)
        clean_val = val_str.replace('.', '').replace(',', '.')
        try:
            val = float(clean_val)
        except ValueError:
            continue
            
        start_idx = max(0, match.start() - 100)
        end_idx = min(len(text), match.end() + 100)
        context = text[start_idx:end_idx].lower()
        
        if (val, val_str) not in seen:
            candidates.append({
                'value': val,
                'str': val_str,
                'context': context,
                'pos': match.start()
            })
            seen.add((val, val_str))
    return candidates

def score_candidate(candidate, activity):
    val = candidate['value']
    context = candidate['context']
    score = 0
    act_lower = activity.lower()
    
    # Global exclusions
    if 11600 <= val <= 11999:  # CEPs in São Sebastião / Litoral Norte
        score -= 120
    if val > 100000:  # Likely phone number or CNPJ
        score -= 200
    if val in [2024.0, 2025.0, 2026.0, 2027.0]:  # Current years
        score -= 80
        
    # Heuristics based on activity
    if 'hospedagem' in act_lower or 'hotel' in act_lower or 'pousada' in act_lower or 'hostel' in act_lower or 'camping' in act_lower:
        if val < 40:  # Too cheap for a room rate
            score -= 50
        elif 80 <= val <= 3000:
            score += 20
            
        keywords = ['diária', 'diaria', 'casal', 'quarto', 'apartamento', 'apto', 'tarifa', 'suíte', 'suite', 'a partir', 'diárias']
        for kw in keywords:
            if kw in context:
                score += 15
                
        # Proximity boost
        if re.search(r'(?:a partir de|diária de|diaria de|diárias|diarias|suíte|suite)\s*(?:r\$\s*)?' + re.escape(candidate['str']), context):
            score += 30
            
        negatives = ['telefone', 'cep', 'cnpj', 'whatsapp', 'tel', 'contato', 'fone', 'criança', 'crianca']
        for neg in negatives:
            if neg in context:
                score -= 12
                
    elif 'restaurante' in act_lower or 'alimentação' in act_lower or 'bar' in act_lower or 'similares' in act_lower:
        if val < 10:  # Too cheap for a meal
            score -= 30
        elif val > 400:  # Too expensive for a restaurant meal
            score -= 60
        elif 15 <= val <= 250:
            score += 20
            
        keywords = ['prato', 'almoço', 'almoco', 'jantar', 'porção', 'porcao', 'refeição', 'refeicao', 'cardápio', 'cardapio', 'menu', 'pf', 'buffet', 'rodízio', 'rodizio', 'individual', 'pizza']
        for kw in keywords:
            if kw in context:
                score += 15
                
        if re.search(r'(?:prato|almoço|almoco|refeição|porção|individual|rodízio)\s*(?:r\$\s*)?' + re.escape(candidate['str']), context):
            score += 25
            
    else:  # Other generic activities (Tours, locadoras, boat rental, etc.)
        if val < 10:
            score -= 50
        elif 25 <= val <= 1000:
            score += 20
            
        keywords = ['passeio', 'ingresso', 'tour', 'embarcação', 'lancha', 'barco', 'roteiro', 'mergulho', 'pessoa', 'adulto', 'valor', 'preço', 'preco', 'tarifa', 'diária', 'diaria', 'locação', 'locacao', 'aluguel']
        for kw in keywords:
            if kw in context:
                score += 15
                
        if re.search(r'(?:por pessoa|ingresso|passeio|lancha|rapel|trilha|diária|diaria|locação|aluguel)\s*(?:r\$\s*)?' + re.escape(candidate['str']), context):
            score += 20
            
    return score

def scrape_establishment_price(url, activity):
    try:
        # 1. Fetch homepage
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, verify=False)
        if resp.status_code != 200:
            return None, f"Erro HTTP {resp.status_code}"
            
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        # 2. Extract potential internal links for deeper crawling if homepage yields nothing
        links_to_crawl = []
        keywords = ['acomodac', 'quarto', 'chale', 'reserva', 'tarifa', 'preco', 'valor', 'suite', 'hosped', 'tarifas', 'cardapio', 'menu']
        for a in soup.find_all('a', href=True):
            href = a['href']
            text = a.get_text().lower()
            href_lower = href.lower()
            matches_kw = any(kw in href_lower or kw in text for kw in keywords)
            if matches_kw and is_internal(url, href):
                full_url = urljoin(url, href)
                if full_url not in links_to_crawl and full_url != url and not '#' in href:
                    links_to_crawl.append(full_url)
                    
        # 3. Collect page text contents (Homepage + up to 2 subpages)
        pages_text = [clean_text(resp.text)]
        for sub_url in links_to_crawl[:2]:
            try:
                sub_resp = requests.get(sub_url, headers=HEADERS, timeout=6, verify=False)
                if sub_resp.status_code == 200:
                    pages_text.append(clean_text(sub_resp.text))
            except Exception:
                pass
                
        # 4. Search and score candidates
        all_candidates = []
        for text in pages_text:
            candidates = extract_price_candidates(text)
            for c in candidates:
                score = score_candidate(c, activity)
                all_candidates.append((score, c))
                
        if not all_candidates:
            return None, "Preço não encontrado"
            
        # Sort candidates: highest score first, then lowest value (preferring starting/starting rates)
        all_candidates.sort(key=lambda x: (-x[0], x[1]['value']))
        
        best_score, best_cand = all_candidates[0]
        if best_score < 0:
            return None, f"Preço de baixa relevância (Score: {best_score})"
            
        return best_cand['value'], "Sucesso"
        
    except requests.exceptions.Timeout:
        return None, "Erro: Timeout de conexão"
    except requests.exceptions.ConnectionError:
        return None, "Erro: Falha de conexão/DNS"
    except Exception as e:
        return None, f"Erro: {str(e)}"

def run_monitor():
    if not os.path.exists(CSV_PATH):
        print(f"Erro: Arquivo {CSV_PATH} não encontrado na pasta atual.")
        return
        
    # Read rows
    with open(CSV_PATH, mode='r', encoding='utf-8') as f:
        reader = csv.reader(f)
        header = next(reader)
        rows = list(reader)
        
    # Standardize header positions
    cols = {name: idx for idx, name in enumerate(header)}
    
    # Ensure missing columns are present
    required_cols = ['Preco Anterior', 'Preco Atual', 'Variacao', 'Status', 'Ultima Atualizacao']
    for rc in required_cols:
        if rc not in cols:
            header.append(rc)
            for r in rows:
                r.append('')
            cols[rc] = len(header) - 1
            
    print(f"Carregados {len(rows)} estabelecimentos do arquivo {CSV_PATH}.")
    print("Iniciando monitoramento de preços (Sem IA - 100% Gratuito)...")
    print("="*80)
    
    success_count = 0
    not_found_count = 0
    error_count = 0
    
    # Use tqdm progress bar if available
    try:
        from tqdm import tqdm
        iterator = tqdm(enumerate(rows), total=len(rows), desc="Progresso")
    except ImportError:
        iterator = enumerate(rows)
        
    for idx, row in iterator:
        id_est = row[cols['ID']]
        nome = row[cols['Nome Fantasia']]
        activity = row[cols['Atividade']]
        url = row[cols['Site']].strip()
        
        # Skip invalid URLs
        if not url or url == '-' or 'instagram.com' in url or 'facebook.com' in url or '@' in url:
            row[cols['Status']] = "Plataforma JS ou URL Inválida"
            row[cols['Ultima Atualizacao']] = time.strftime('%d/%m/%Y %H:%M')
            continue
            
        # Handle prefix fixes (e.g. ww. or missing http)
        if url.startswith('ww.'):
            url = 'https://www.' + url[3:]
        elif not url.startswith('http'):
            url = 'https://' + url
            
        # Scrape price
        preco_anterior = row[cols['Preco Atual']]
        novo_preco, status = scrape_establishment_price(url, activity)
        
        # Update row data
        row[cols['Status']] = status
        row[cols['Ultima Atualizacao']] = time.strftime('%d/%m/%Y %H:%M')
        
        if novo_preco is not None:
            row[cols['Preco Anterior']] = preco_anterior if preco_anterior else ''
            row[cols['Preco Atual']] = f"{novo_preco:.2f}".replace('.', ',')
            
            # Calculate variation
            if preco_anterior:
                try:
                    prev_float = float(preco_anterior.replace(',', '.'))
                    if prev_float > 0:
                        var = (novo_preco - prev_float) / prev_float
                        row[cols['Variacao']] = f"{var:.2%}".replace('.', ',')
                except ValueError:
                    row[cols['Variacao']] = ''
            else:
                row[cols['Variacao']] = ''
                
            success_count += 1
            if not isinstance(iterator, list) and 'tqdm' in sys.modules:
                pass # tqdm automatically prints
            else:
                print(f"[✅ Sucesso] {nome}: R$ {novo_preco:.2f}")
        else:
            if "Erro" in status:
                error_count += 1
            else:
                not_found_count += 1
            if not isinstance(iterator, list) and 'tqdm' in sys.modules:
                pass
            else:
                print(f"[❌ {status}] {nome}")
                
        # Save back to CSV in real-time to avoid data loss on interrupt
        with open(CSV_PATH, mode='w', newline='', encoding='utf-8') as fw:
            writer = csv.writer(fw)
            writer.writerow(header)
            writer.writerows(rows)
            
        # Delay between requests
        time.sleep(DELAY_BETWEEN_SITES)
        
    print("="*80)
    print("Monitoramento Concluído!")
    print(f"Planilha {CSV_PATH} atualizada com sucesso.")
    print(f"✅ Preço encontrado    : {success_count}")
    print(f"🔍 Preço não encontrado: {not_found_count}")
    print(f"❌ Erros de conexão     : {error_count}")

if __name__ == "__main__":
    run_monitor()
