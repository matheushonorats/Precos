# Sistema de Monitoramento de Preços Turísticos — SETUR

Este projeto consiste em um sistema automatizado para monitorar e extrair preços em tempo real dos estabelecimentos turísticos de São Sebastião/SP que possuem site na internet (351 estabelecimentos identificados no cadastro original).

O sistema oferece duas abordagens complementares de implementação:
1. **Google Sheets + Apps Script (com IA - Recomendado)**: Roda diretamente na nuvem no Google Planilhas integrado à API do Gemini para interpretar páginas complexas e dinâmicas de forma semântica.
2. **Python Local (Sem IA - 100% Gratuito)**: Um script em Python para execução local que utiliza web scraping tradicional (BeautifulSoup) e heurísticas avançadas de pontuação de texto para encontrar os preços na estrutura do site.

---

## 📋 Especificações Técnicas

### 1. Dados e Estrutura do CSV (`estabelecimentos_para_monitorar.csv`)
O arquivo de cadastro contém os estabelecimentos contendo os seguintes campos:
* **ID**: Identificador exclusivo do cadastro (ex: `SS0589`).
* **Nome Fantasia**: Nome comercial do estabelecimento.
* **Atividade**: Atividade turística principal (ex: Hospedagem, Restaurantes, Locadoras).
* **Site**: Endereço web do estabelecimento (utilizado para a captura dos preços).
* **Preço Anterior**: Preço registrado no ciclo anterior de monitoramento.
* **Preço Atual**: Preço mais recente extraído.
* **Variação**: Variação percentual entre o preço atual e o anterior.
* **Status**: Status da captura (ex: `Sucesso`, `Erro HTTP 404`, `Preço não encontrado`, `Plataforma JS ou URL Inválida`).
* **Última Atualização**: Data e hora da última tentativa de scraping.

---

## 🛠️ Modos de Execução

### Opção A: Google Sheets + Apps Script (Gemini API)
Implementado no arquivo [CodigoMonitoramento_v2.js](file:///C:/Users/mathe/Desktop/Preços/CodigoMonitoramento_v2.js).

* **Como Funciona**: 
  1. O script faz o download do HTML do site usando o `UrlFetchApp`.
  2. Executa um algoritmo de limpeza profunda que remove scripts, estilos (CSS), SVGs, comentários e tags HTML, reduzindo em média **99% do payload**.
  3. Envia o texto comprimido e o prompt com regras específicas de negócio para a **Gemini API (Modelo `gemini-3.1-flash-lite` ou `gemini-1.5-flash`)**.
  4. O Gemini localiza o preço com base no contexto, retornando um JSON formatado com o preço, moeda e descrição do que foi coletado.
* **Recursos Avançados da Versão 2.0**:
  - **Processamento em Lotes (Queue System)**: Contorna o limite máximo de execução de 6 minutos do Google Apps Script salvando o estado e permitindo retomar de onde parou.
  - **Gatilhos Automáticos (Triggers)**: Agendamento diário automático para execução na madrugada (ex: 2h).
  - **Abas Auxiliares**: Gerenciamento automático de histórico de preços de auditoria e log detalhado de falhas/erros HTTP.
  - **Detecção de Plataformas Dinâmicas**: Reconhece se o site pertence a redes sociais ou sistemas de agendamento fechados (ex: Instagram, iFood, stays.net) e ajusta o status sem gastar chamadas de API desnecessárias.

### Opção B: Script Local em Python
Implementado no arquivo [monitor_precos.py](file:///C:/Users/mathe/Desktop/Preços/monitor_precos.py).

* **Como Funciona**:
  1. Realiza requisições HTTP (com `requests`) usando User-Agent simulado e lidando com certificados SSL legados.
  2. Extrai links internos de páginas chaves (como "Acomodações", "Tarifas", "Cardápio") para vasculhar até duas subpáginas além da Home.
  3. Localiza valores monetários no formato brasileiro (`R$ X.XXX,XX`) através de Regex.
  4. Executa um **algoritmo de pontuação e heurística (Scoring)** baseado nas regras de negócio para eleger o preço mais provável.
* **Regras de Negócio do Scoring**:
  - **Exclusões Globais**: Ignora anos atuais (ex: `2024`, `2025`, `2026`), faixas de CEP de São Sebastião (`11600` a `11999`), números telefônicos ou CNPJs.
  - **Hospedagem**: Penaliza preços excessivamente baixos (< R$ 40,00). Dá bônus de relevância para termos como `diária`, `casal`, `tarifa`, `suíte` e proximidade de frases como `"a partir de"`.
  - **Restaurantes**: Penaliza preços abaixo de R$ 10,00 ou acima de R$ 400,00. Pontua positivamente pratos principais, almoço, jantar, pizzas, cardápio, etc.
  - **Outros/Genéricos**: Pontua passeios, ingressos, aluguéis de embarcações, diárias de locação de automóvel básico.

---

## ⚙️ Regras de Negócio para Coleta de Preços

O sistema foi modelado para extrair o valor mais representativo de cada atividade de acordo com as seguintes diretrizes:

| Segmento | Preço Alvo a Ser Capturado |
| :--- | :--- |
| **Hospedagem / Acampamento** | Tarifa de diária inicial, diária mínima ou tarifa "a partir de" para quarto casal/duplo padrão. |
| **Restaurantes / Alimentação** | Preço médio de prato principal individual, prato do dia (PF), buffet, rodízio ou pizza básica. |
| **Locadoras de Veículos** | Valor da diária padrão para aluguel de veículo econômico básico. |
| **Turismo Náutico / Passeios** | Preço por pessoa para o passeio ou roteiro básico ofertado na Home. |

---

## 🚀 Como Iniciar

Consulte o arquivo de orientações para o guia detalhado de configuração da planilha Google e ativação do painel de controle.

Para executar o script local em Python:
1. Certifique-se de que os pacotes necessários estejam instalados:
   ```bash
   pip install requests beautifulsoup4 urllib3
   ```
2. Execute o monitor:
   ```bash
   python monitor_precos.py
   ```
3. O progresso será exibido no console e os dados do arquivo `estabelecimentos_para_monitorar.csv` serão atualizados incrementalmente em tempo real.
