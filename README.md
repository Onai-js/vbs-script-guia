[![wiki](https://img.shields.io/static/v1?label=VBS&message=v5.8&color=blueviolet)](https://pt.wikipedia.org/wiki/VBScript) 

# VBScript
O VBScript é frequentemente usado em substituição aos arquivos de lote, sendo uma opção mais avançada e prática

VBScript usa o Component Object Model para acessar os elementos do ambiente no qual ele está sendo executado, o que permite o uso de funções avançadas do sistema operacional, por exemplo, para manipulação de arquivos ou do registro do Windows.


**OBSERVAÇÃO:** Geralmente maioria dos sites que tentamos fazer um up de um arquivo .vbs reconhecem como vírus só pelo fato de ter a extensão .vbs então pode ser um pouco mais dificil de enviar esses projetos

## Acessos rápidos

* [Guia / Comece](https://riptutorial.com/vbscript) _(guia não atualizado)_
* [GitHub](https://github.com/Onai-js/vbs-script-guia)
* [Wikipédia](https://pt.wikipedia.org/wiki/VBScript)

## No editor de códigos fontes...

Aplique um nome que possua .vbs no final do nome para
ser reconhecido como um arquivo VBScript.

## Examplo de um código

```vbs

x=msgbox ("Apenas" ,0, "Exemplo") 

dim result 

result = msgbox("Ta bom?", 4 , "Confirmação") 'Select, de SIM ou NÃO

If result=6 then ' Checando se a resposta foi "Sim".

msgbox "Sim foi selecionado"

else 'Caso tenha sido "Não"

msgbox "Não foi selecionado"

end if ' Fechando o IF

```

Você viu que possui um número e toda declaração de caixa de mensagem? que no caso da primeira declaração ```x=msgbox ("Apenas" ,0, "Exemplo")``` é esse 0(zero) que fica após a vírgula de "Apenas".


Isso serve pra declarar qual tipo de caixa de mensagem vai ser, no caso qual botão ou input vai aparecer. Siga a tabela abaixo que indica o que cada número é


## Tabela de botões

| Números  | Valores |
| ------------- | ------------- |
| 0 = vbOKOnly  | apenas botão OK  |
| 1 = vbOKCancel  | OK e Cancelar botões  |
| 2 = vbAbortRetryIgnore  | Abortar, tentar e ignorar botões  |
| 3 = vbYesNoCancel  | Sim, Não, e Cancelar botões  |
| 4 = vbYesNo  | Botões Sim e Não  |
| 5 = vbRetryCancel | Botões de retimento e cancelamento |
| 16 = vbCritical  | Ícone da Mensagem Crítica  |
| 32 = vbQuestion | Ícone de consulta de aviso |
| 48 = vbExclamation | Ícone da mensagem de aviso |
| 64 = vbInformation | Ícone da Mensagem de Informação |
| 0 = vbDefaultButton1 | O primeiro botão é padrão |
| 256 = vbDefaultButton2 | O segundo botão é padrão | 
| 512 = vbDefaultButton3 | Terceiro botão é padrão |
| 768 = vbDefaultButton4  | Quarto botão é padrão |
| 0 = vbApplicationModal  | Aplicativo modal (o aplicativo atual não funcionará até que o usuário responda à caixa de mensagens) |
| 4096 = vbSystemModal  | Modal do sistema (todos os aplicativos não fucionarão até que o usuário responda á caixa de mensagens)  |

## Isenção de responsabilidade

Todas as informações contidas aqui, são frutos de uma vasta pesquisa e conhecimentos básicos. Embora não seja patrocinada ou remunerada pelo [GitHub](https://github.com) ou [Microsoft](https://www.microsoft.com/pt-br/). Foi feita apenas á fins educativos e práticos!

## Licensa desse documento

Copyright 2022 Onaissac

