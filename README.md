# CellCaster

**Um motor de raycasting 3D com texturas rodando dentro do Excel. Cada pixel é uma célula da planilha. VBA puro, sem suplementos, sem DLLs, sem arquivos binários.**

[Read in English](README.en.md)

![Dois quadros renderizados pelo CellCaster](docs/preview.png)

> A imagem acima é desenhada por `tools/reference_renderer.py`, uma porta em Python verificada para produzir exatamente as mesmas células que o VBA. Troque por um GIF gravado no Excel quando tiver um.

Ande por um labirinto no estilo Wolfenstein, colete as seis gemas e bata o seu tempo, tudo em um bloco de 128 x 72 células.

## Situação

Jogável, mas lento: na primeira máquina em que foi testado (Excel para Windows, 128 x 72 células) roda a 2-3 quadros por segundo. O HUD mostra quanto de cada quadro vai para o desenho, para a escrita na planilha e para a repintura do Excel. Deixar mais rápido é o primeiro item do roteiro. Por causa da taxa baixa, segure `Esc` por um segundo para sair.

## Recursos

- Paredes com textura, raycasting por DDA e correção de olho de peixe
- Chão quadriculado com floor casting e teto com neblina
- Neblina por distância com dithering ordenado (Bayer 4x4): 4 tons por cor parecem muitos
- Sprites (as gemas) com z-buffer por coluna
- Minimapa, HUD com cronômetro e FPS
- Texturas procedurais: tijolo, pedra, madeira, metal e musgo são geradas por código
- Movimento independente da taxa de quadros, deslizando nas paredes
- O VBA do motor é testado no Linux pelo CI, sem Excel

## Como ele renderiza

Mudar `Interior.Color` em 9.216 células, uma a uma, a cada quadro, é lento demais para um jogo. O CellCaster não mexe em formatação durante o jogo:

1. **Paleta feita de formatação condicional.** Ao iniciar, a área da tela recebe 48 regras do tipo "valor da célula = *n* -> cor *n*" (12 cores x 4 tons; 0 fica preto).
2. **Uma escrita por quadro.** O motor preenche uma matriz 2-D com índices da paleta e grava tudo em uma única chamada a `Range.Value2`. O Excel repinta sozinho.
3. **Dithering em vez de mais cores.** Um limiar de Bayer decide como cada célula arredonda o brilho, então a paleta e o número de regras ficam pequenos.

## Como rodar

Requer Excel para Windows (2010 ou mais novo, 32 ou 64 bits). O jogo lê o teclado pela API do Windows, então não roda no Mac nem no Excel para a web.

**Opção A: script de build**

1. No Excel, habilite *Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Configurações de Macro > Confiar no acesso ao modelo de objeto do projeto do VBA* (uma vez).
2. Rode:
   ```powershell
   powershell -ExecutionPolicy Bypass -File tools\build.ps1
   ```
3. Abra o `CellCaster.xlsm` gerado, habilite as macros e clique em **JOGAR**.

**Opção B: manual**

1. Crie uma pasta de trabalho e salve como `.xlsm`.
2. Aperte `Alt+F11` e arraste os sete arquivos de `src/` para o painel de projeto (ou use *Arquivo > Importar Arquivo*).
3. Aperte `Alt+F8`, execute `InstallMenu` e clique em **JOGAR** (ou execute `StartGame`).

## Controles

| Tecla | Ação |
| --- | --- |
| `W` `S` ou setas para cima e para baixo | Andar |
| `A` `D` | Andar de lado |
| Setas esquerda e direita | Girar |
| `M` | Mostrar ou ocultar o minimapa |
| `Esc` | Sair |

## Estrutura do projeto

| Arquivo | Papel |
| --- | --- |
| `src/Config.bas` | Todas as constantes ajustáveis e a paleta de cores |
| `src/World.bas` | Fase, jogador, colisões, gemas |
| `src/Textures.bas` | Texturas procedurais |
| `src/Raycaster.bas` | Chão, paredes, sprites e minimapa em um buffer de quadro |
| `src/Renderer.bas` | Preparo da planilha e a escrita do quadro em uma chamada |
| `src/Platform.bas` | API do Windows: teclado, relógio, foco |
| `src/Game.bas` | Loop principal, entrada, menu |
| `tools/build.ps1` | Gera o `.xlsm` a partir de `src/` |
| `tools/reference_renderer.py` | Porta em Python do renderizador, validador de fase, imagem de prévia |
| `tests/` | Executa o VBA de verdade no LibreOffice e compara cada célula |

Os nomes e os comentários do código estão em inglês e sem acentos: o editor do VBA importa arquivos em ANSI, não em UTF-8, então qualquer acento apareceria quebrado no Excel ou no GitHub.

## Crie a sua fase

Edite as linhas `AddRow` em `src/World.bas`. `1`-`5` são texturas de parede, `.` é chão, `P` é o início do jogador e `*` é uma gema. A borda precisa ser toda de parede. Depois valide:

```bash
python tools/reference_renderer.py --check
```

## Testes

`Config`, `World`, `Textures` e `Raycaster` não chamam o Excel nem o Windows, então rodam sem alteração no modo VBA do LibreOffice Basic. A suíte executa esses módulos lá, renderiza vários quadros e compara as 9.216 células de cada um com a porta de referência em Python, além de uma caminhada com colisões e da coleta das gemas.

```bash
python tests/run_vba_tests.py    # precisa do LibreOffice (soffice) no PATH
```

`Platform`, `Renderer` e `Game` dependem do Excel no Windows e não são cobertos pelos testes automáticos.

## Roteiro

- Desempenho: menos regras de paleta, tela padrão menor, atualizações parciais
- Portas e uma saída da fase
- Inimigos e um segundo tipo de sprite
- Olhar com o mouse
- Script de exportação para levar ao `src/` o que for editado no editor do VBA
- Editor de fases em uma planilha: pinte o mapa com células

## Licença

[MIT](LICENSE)
