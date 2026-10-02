# Conferência XML Fiscal

Aplicação web para processar XML de NF-e e NFC-e, analisar o lote e exportar uma planilha em Excel com resumo, itens, totais por CFOP e conferência de problemas comuns.

## O que o projeto faz

- aceita arquivos .xml e .zip
- lê notas fiscais diretamente no navegador
- valida dados de notas, itens e impostos
- identifica numeração pulada, duplicadas e canceladas
- gera uma planilha .xlsx pronta para conferência e trabalho manual
- mantém o processamento local, sem envio de documentos para servidor

## Público alvo

- contadores
- departamentos fiscais
- pequenas empresas que recebem ou emitem lotes de notas
- quem precisa conferir XMLs rapidamente antes de exportar dados

## Stack

- Vite
- JavaScript vanilla
- JSZip
- ExcelJS
- HTML/CSS

## Como rodar localmente

```bash
npm install
npm run dev
```

Acesse a URL exibida no terminal pelo Vite.

## Como gerar build de produção

```bash
npm run build
```

Os arquivos prontos ficam na pasta dist.

## Como publicar

Você pode publicar a pasta dist em plataformas como Cloudflare Pages, Netlify ou Vercel.

### Cloudflare Pages

1. conecte o repositório
2. use o comando de build: npm run build
3. defina a pasta de saída: dist
4. publique

## Observação importante

A ferramenta é útil como apoio de conferência e organização de dados fiscais, mas não substitui revisão humana e critérios próprios do usuário ou do contador responsável.

## Melhorias futuras sugeridas

- upload de arquivos CSV e XLSX para cruzamento de dados
- comparação entre meses
- exportação em PDF
- dashboard com indicadores resumidos
- autenticação para uso em equipe
- versão premium com geração de relatórios executivos