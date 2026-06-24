# BRDrive Dashboard Web â€” J&T Express SP

Gerador de dashboard HTML completo a partir dos dados exportados do sistema de rastreamento BRDrive. Produz um arquivo HTML standalone com 15 pÃ¡ginas analÃ­ticas que pode ser compartilhado internamente ou servido via servidor local.

---

## Funcionalidades

- GeraÃ§Ã£o de dashboard HTML com 15 pÃ¡ginas analÃ­ticas a partir de Excel/CSV
- Arquivo de saÃ­da standalone â€” nÃ£o precisa de backend para visualizaÃ§Ã£o
- Servidor local para compartilhamento na rede interna
- Script de atualizaÃ§Ã£o automÃ¡tica com bat
- Fixa IP local para acesso consistente via rede
- CompatÃ­vel com mÃºltiplas mÃ¡quinas (resolve automaticamente o caminho dos dados)

## Tecnologias

| Componente | Tecnologia |
|------------|-----------|
| GeraÃ§Ã£o | Python, Pandas, NumPy |
| VisualizaÃ§Ãµes | HTML5, CSS3, JavaScript |
| Dados | Excel / CSV (exportaÃ§Ã£o BRDrive) |
| Servidor | Python HTTP server |
| AutomaÃ§Ã£o | .bat (Windows) |

## Estrutura

```
BRDrive_Dash_Web/
â”œâ”€â”€ dashboard.py          # Gerador principal (15 pÃ¡ginas)
â”œâ”€â”€ server.py             # Servidor HTTP para rede local
â”œâ”€â”€ servir.py             # Alternativa de servidor
â”œâ”€â”€ brdrive_output/       # HTMLs gerados
â”œâ”€â”€ atualizar_dashboard.bat  # Atualiza e reabre o dashboard
â”œâ”€â”€ iniciar_servidor.bat  # Inicia servidor local
â””â”€â”€ fixar_ip.bat          # Configura IP fixo local
```

## Como usar

```bash
# 1. Instalar dependÃªncias
pip install pandas numpy

# 2. Atualizar o dashboard (apÃ³s exportar dados do BRDrive)
atualizar_dashboard.bat

# 3. Ou iniciar servidor para compartilhar na rede
iniciar_servidor.bat
```

## Contexto

O BRDrive nÃ£o oferece dashboards customizados. Este projeto gerou uma camada analÃ­tica prÃ³pria sobre os dados exportados, com visualizaÃ§Ãµes adaptadas Ã s necessidades da operaÃ§Ã£o SP, e entregou os resultados de forma que qualquer pessoa com acesso Ã  rede pudesse consultar sem precisar de instalaÃ§Ã£o.

---

*Desenvolvido por Robson Noberto â€” Analista de Dados | J&T Express Filial SP*

