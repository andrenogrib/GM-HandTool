# GM-HandTool
![alt tag](https://i.imgur.com/kbxjaMY.jpg"")
![alt tag](https://i.imgur.com/Cfvzb7w.jpg"")
![alt tag](https://i.imgur.com/zkxGw5p.jpg"")
![alt tag](https://i.imgur.com/JbuHJa0.jpg"")
![alt tag](https://i.imgur.com/89rkXTs.jpg"")
![alt tag](https://i.imgur.com/h5S7PgJ.jpg"")

Aplicativo WinForms em PascalABC.NET para ler arquivos `.wz` do MapleStory, visualizar dados por aba, salvar em BIN e exportar para uma pagina web pesquisavel.

## Funcionalidades
- Leitura direta de WZ (`Load From WZ`).
- Leitura de BIN salvo anteriormente (`Load From BIN`).
- Busca textual na aba atual.
- Exportacao BIN por aba (`Save BIN`).
- Exportacao WEB (`Export WEB`) com JSON + imagens + pagina de busca.

## Requisitos
- Windows
- PascalABC.NET (IDE ou compilador de linha de comando)
- Arquivos WZ do cliente MapleStory (ex.: `Item.wz`, `Character.wz`, `String.wz` etc.)

## Compilacao

### Opcao 1: IDE
1. Abra `GMHandTool/GMHandTool.pabcproj` no PascalABC.NET.
2. Compile/Run pelo proprio IDE.

### Opcao 2: terminal
Exemplo:

```powershell
"C:\Program Files (x86)\PascalABC.NET\pabcnetcclear.exe" ".\GMHandTool\GMHandTool.pabcproj"
```

## Como usar (desktop)
1. Abra o app `GMHandTool.exe`.
2. Escolha `Load From WZ`.
3. Clique no campo de pasta e selecione a pasta do MapleStory.
4. Selecione a aba desejada (Cash, Consume, Weapon, etc.).
5. Clique `Load WZ`.

Observacao:
- O app foi ajustado para resolver `name/description` em estruturas de `String.wz` com caminhos como `Item.img/Con`, `Item.img/Ins`, `Item.img/Eqp/...`.

## BIN
- `Save BIN` salva apenas a aba atual em `<TabName>.BIN` no diretorio atual.
- `Load From BIN` abre os BINs previamente salvos por aba.

## Exportacao WEB
O botao `Export WEB` exporta todas as abas que ja possuem linhas carregadas no momento.

Fluxo:
1. Carregue as abas que voce quer publicar.
2. Clique `Export WEB`.
3. Abra `GMHandTool/WEB/index.html` no navegador.

Saida gerada:
- `WEB/<TabName>/data.json`
- `WEB/<TabName>/data.js`
- `WEB/<TabName>/images/*.png`
- `WEB/manifest.js`
- `WEB/index.html`

## Estrutura do projeto
- `GMHandTool/MainUnit.pas`: logica principal de carregamento das abas, busca e eventos de UI.
- `GMHandTool/DataGrid.pas`: tipos de grid, BIN e exportacao WEB.
- `GMHandTool/WzUtils.pas`: helpers/extensoes para navegar em WZ e extrair PNG.
- `GMHandTool/MainUnit.MainForm.inc`: definicao dos componentes WinForms.

## Creditos
- WZ library: `WZComparerR2`
- Compiler/IDE: `PascalABC.NET`
  - http://pascalabc.net/en/
