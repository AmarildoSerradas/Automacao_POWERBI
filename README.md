# 📊 Automação de Relatórios Power BI + WhatsApp

Script de automação desenvolvido em Python para captura, processamento e envio de relatórios do Power BI de forma totalmente automatizada.

O sistema acessa dashboards online, aplica filtros dinâmicos (como datas e lojas), captura screenshots dos relatórios, realiza conversão de imagens e envia automaticamente para grupos no WhatsApp Web.

# ⚙️ Tecnologias utilizadas
Python
Playwright (automação web)
Pillow (processamento de imagens)
win32clipboard (integração com área de transferência)
# 🚀 Funcionalidades
Acesso automático a múltiplos relatórios do Power BI
Aplicação de filtros por data e loja
Captura de screenshots dos dashboards
Conversão automática de imagens para JPG
Organização em pastas por tipo de relatório
Envio automático das imagens via WhatsApp Web
Execução contínua sem intervenção manual
# 🧠 Como funciona

O script utiliza o Playwright para simular a navegação do usuário, acessando relatórios no Power BI, aplicando filtros e capturando os dados visuais.

Após isso, as imagens são tratadas e enviadas automaticamente para grupos específicos no WhatsApp.

# ▶️ Como executar
pip install playwright pillow pywin32
playwright install
python nome_do_arquivo.py
# ⚠️ Pontos de atenção
Dependente de elementos da interface (pode quebrar se o Power BI mudar)
Uso de time.sleep() ao invés de waits inteligentes
Código grande e pouco modular
Dependência do WhatsApp Web logado manualmente
# 📌 Melhorias futuras
Refatoração com funções e organização por módulos
Implementação de waits inteligentes no Playwright
Logs de execução e tratamento de erros mais robusto
Agendamento automático (Task Scheduler / cron)
