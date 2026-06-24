# BRDrive Dashboard Web — J&T Express SP

Gerador de dashboard HTML completo a partir dos dados exportados do sistema de rastreamento BRDrive. Produz um arquivo HTML standalone com 15 páginas analíticas que pode ser compartilhado internamente ou servido via servidor local.

---

## Funcionalidades

- Geração de dashboard HTML com 15 páginas analíticas a partir de Excel/CSV
- Arquivo de saída standalone — não precisa de backend para visualização
- Servidor local para compartilhamento na rede interna
- Script de atualização automática com bat
- Fixa IP local para acesso consistente via rede
- Compatível com múltiplas máquinas (resolve automaticamente o caminho dos dados)

## Tecnologias

| Componente | Tecnologia |
|------------|-----------|
| Geração | Python, Pandas, NumPy |
| Visualizações | HTML5, CSS3, JavaScript |
| Dados | Excel / CSV (exportação BRDrive) |
| Servidor | Python HTTP server |
| Automação | .bat (Windows) |

## Estrutura

```
BRDrive_Dash_Web/
├── dashboard.py          # Gerador principal (15 páginas)
├── server.py             # Servidor HTTP para rede local
├── servir.py             # Alternativa de servidor
├── brdrive_output/       # HTMLs gerados
├── atualizar_dashboard.bat  # Atualiza e reabre o dashboard
├── iniciar_servidor.bat  # Inicia servidor local
└── fixar_ip.bat          # Configura IP fixo local
```

## Como usar

```bash
# 1. Instalar dependências
pip install pandas numpy

# 2. Atualizar o dashboard (após exportar dados do BRDrive)
atualizar_dashboard.bat

# 3. Ou iniciar servidor para compartilhar na rede
iniciar_servidor.bat
```

## Contexto

O BRDrive não oferece dashboards customizados. Este projeto gerou uma camada analítica própria sobre os dados exportados, com visualizações adaptadas às necessidades da operação SP, e entregou os resultados de forma que qualquer pessoa com acesso à rede pudesse consultar sem precisar de instalação.

---

*Desenvolvido por Robson Noberto — Analista de Processos | J&T Express Filial SP*
