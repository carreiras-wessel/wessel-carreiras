# Wessel Carreiras

Portal de descritivos de cargos da Wessel. O site é estático e publicado no
GitHub Pages. A **fonte oficial dos dados é a planilha** em
`dados/Base_de_Cargos_Wessel.xlsx`; o portal é gerado a partir dela, nunca
editado à mão.

## Como atualizar os cargos (fluxo do dia a dia)

1. Abra `dados/Base_de_Cargos_Wessel.xlsx` no Excel e faça as alterações.
2. Salve o arquivo com o **mesmo nome**.
3. Envie a planilha atualizada para o repositório, na pasta `dados/`
   (pelo site do GitHub: entre na pasta `dados`, clique em **Add file >
   Upload files**, arraste a planilha por cima da existente e confirme).
4. Pronto. O GitHub Actions detecta a mudança, reconstrói o site e publica
   sozinho em cerca de 1 a 2 minutos. Não é preciso mexer em código.

O endereço público do portal aparece em **Settings > Pages** depois da
primeira publicação.

## O que o portal faz

- Navegação por **Área** e por **Setor**.
- **Busca** por título, código, setor ou atividade.
- **Descritivo completo** de cada cargo (identificação, missão,
  responsabilidades, requisitos, competências, controle).
- Campos ainda não levantados aparecem como **"a definir"**, nunca inventados.
- **Exportação** de um cargo individual ou da base completa (com os filtros
  aplicados) em Excel.

## Estrutura do repositório

```
dados/     planilha oficial (você edita aqui)
scripts/   build.py  -> converte a planilha em site/data.js
site/      o portal publicado (index.html + assets)
.github/   automação de publicação (GitHub Actions)
```

## Rodar e testar na sua máquina (opcional)

Precisa de Python instalado.

```bash
pip install openpyxl
python scripts/build.py
python -m http.server 8000 --directory site
```

Depois abra `http://127.0.0.1:8000` no navegador.

## Campos que ainda serão preenchidos

Nível/Grau, Cargo Superior Imediato, Subordinados, Missão do Cargo,
Competências Técnicas, Treinamentos/NR, Requisitos Físicos/Ambiente e
Viagens/CNH não existiam no arquivo original. Ficam como "a definir" até
serem mapeados com os gestores. Basta preencher na planilha e reenviar; o
portal passa a exibir o conteúdo.

## Fase 2 (planejado)

Controle de acesso por senha para limitar a visualização. Os dados já ficam
num arquivo separado (`site/data.js`) justamente para receber a criptografia
sem retrabalho.
