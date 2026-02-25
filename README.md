# logistics-rpa-canhoto-automation
Automação (RPA) para coleta, conversão e organização de canhotos logísticos de múltiplos portais usando Python e Playwright.

## Visão Geral
Este projeto foi desenvolvido para resolver um desafio crítico no setor logístico: a coleta manual e organização de canhotos de entrega (Proof of Delivery - POD). Utilizando Python e RPA, a solução automatiza o acesso a múltiplos portais de transportadoras, realiza o download de documentos, converte-os para imagem e padroniza a nomenclatura para faturamento.

## Tecnologias Utilizadas
* `Python` 3.10+
* `Playwright`: Automação de navegador e Web Scraping.
* `Pandas`: Manipulação de dados e leitura da planilha mestra.
* `PyMuPDF` (fitz): Processamento e conversão de documentos PDF para JPG.
* `Shutil/OS`: Gerenciamento automatizado de arquivos e diretórios.

## Funcionalidades
* `Integração com Planilha Mestra`: Lê automaticamente os dados de NF, Pedido e Cliente.
* `Navegação Inteligente`: Realiza login e busca automatizada em portais ESL Cloud e Brudam.
* `Tratamento de Imagens`: Converte canhotos PDF em JPG com 150 DPI para garantir legibilidade.
* `Padronização`: Organiza os arquivos em pastas por transportadora com nomes padronizados: Canhoto_NF_Pedido_Cliente.jpg.

## Estrutura do Projeto
* `/dados`: Local para a planilha de entrada.
* `/saida`: Diretório onde os canhotos processados são organizados.
* `rpa_canhoto.py`: Script principal da automação.

## Nota de Segurança: 
Por questões de privacidade, todas as URLs e credenciais de acesso foram substituídas por placeholders genéricos.

## Escalabilidade e Próximos Passos
Para operações de grande volume, o projeto foi estruturado pensando em futuras expansões:

* `Integração com Bancos de Dados`: Substituição da planilha Excel por uma conexão direta com bancos SQL (PostgreSQL/SQL Server) para consultas em tempo real.
* `Processamento em Paralelo`: Implementação de Multiprocessing para que o script acesse múltiplas transportadoras simultaneamente, reduzindo o tempo total de execução.
* `Notificações Automáticas`: Integração com APIs de comunicação (Telegram/Slack) para enviar o relatório de pendências (relatorio_geral) automaticamente ao fim do processo.
* `Dashboard de Auditoria`: Conexão dos logs de sucesso e erro com Power BI para visualização do KPI de "Taxa de Comprovação de Entrega".
* `Containerização`: Uso de Docker para garantir que o RPA rode em qualquer servidor ou ambiente de nuvem sem conflitos de dependências.
