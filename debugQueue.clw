  MEMBER

  MAP
  END

  INCLUDE( 'debugQueue.inc' ), ONCE

phDebugQueue            PROCEDURE( QUEUE pQ, <STRING pFile>, BYTE pFormat=1, <STRING pFieldSeparator>, <STRING pStringDelimiter>, BYTE pRecNumbers=TRUE )
lColumnWidth            LONG,AUTO                           !Largura calculada das colunas
zstrColuna              CSTRING( 50 )                       !Nome da coluna
pComponent              ANY                                 !Ponteiro para um componente da queue
lLastX                  LONG,STATIC                         !Última posição X
lLastY                  LONG,STATIC                         !última posição Y
lSavPointer             LONG,AUTO                           !Posição original do ponteiro da queue
bignoredebugqueue       BYTE,STATIC                         !Ignorar a funcionalidade da rotina

phDebugQueueWindow      WINDOW('Debug Queue'),AT(,,400,219),FONT('Tahoma',8,COLOR:Black,FONT:regular,CHARSET:ANSI), |
                          CENTER,GRAY,RESIZE,MAX,IMM
                          LIST,AT(5,5,390,184),USE(?phDebugQueueListBox1),FLAT,HSCROLL,VSCROLL,FROM(''),ALRT(2)
                          PANEL,AT(5,194,390,1),USE(?phDebugQueuePanel1),BEVEL(-1,1)
                          BUTTON('&Ignorar'),AT(266,200,40,13),USE(?phDebugQueuebtnIgnorar),FONT('MS Sans Serif',8,, FONT:Regular,CHARSET:ANSI)
                          BUTTON('&Gravar'),AT(310,200,40,13),USE(?phDebugQueuebtnGravar),FONT('MS Sans Serif',8,,FONT:regular,CHARSET:ANSI)
                          BUTTON('&Fechar'),AT(354,200,40,13),USE(?phDebugQueuebtnFechar),LEFT,FONT('MS Sans Serif',8,,FONT:regular,CHARSET:ANSI), |
                             DEFAULT
                        END

  CODE
  IF bignoredebugqueue THEN RETURN END                      !Se está balizado para ignorar, retorna imediatamente
  lSavPointer = POINTER( pQ )                               !Salva o ponteiro atual da queue
  OPEN( phDebugQueueWindow )
  0{ PROP:Text } = 'Debug Queue - ' & RECORDS( pQ ) & ' registros'
  IF lLastX THEN 0{ PROP:XPos } = lLastX END                !Reposiciona a janela, se já havia sido posicionada
  IF lLastY THEN 0{ PROP:YPos } = lLastY END                !em uma chamada anterior.
  ?phDebugQueueListBox1{ PROP:From } = pQ                   ! Utiliza a queue passada como fonte do listbox
  ?phDebugQueueListBox1{ PROP:Column } = TRUE

  DO FormatQueue
  SELECT( ?phDebugQueueListBox1, 1 )
  ACCEPT
    DISPLAY
    CASE EVENT()
    OF EVENT:AlertKey
      CASE KEYCODE()
      OF 2                                                  !2h = MouseRight
        i# = ?phDebugQueueListBox1{ PROPLIST:MouseDownField }
        EXECUTE POPUP( '@D6|@T4|@N12.`2|@S50' )
          ?phDebugQueueListBox1{ PROPLIST:Picture, i# } = '@D6'
          ?phDebugQueueListBox1{ PROPLIST:Picture, i# } = '@T4'
          ?phDebugQueueListBox1{ PROPLIST:Picture, i# } = '@N12.`2'
          ?phDebugQueueListBox1{ PROPLIST:Picture, i# } = '@S50'
        END
      END

    OF EVENT:Sized
      ?phDebugQueueListBox1{ PROP:Width } = 0{ PROP:Width } - 10
      ?phDebugQueueListBox1{ PROP:Height } = 0{ PROP:Height } - 35
      ?phDebugQueuePanel1{ PROP:Width } = ?phDebugQueueListBox1{ PROP:Width }
      ?phDebugQueuePanel1{ PROP:YPos } = 0{ PROP:Height } - 25
      ?phDebugQueuebtnIgnorar{ PROP:XPos } = 0{ PROP:Width } - 134
      ?phDebugQueuebtnIgnorar{ PROP:YPos } = 0{ PROP:Height } - 19
      ?phDebugQueuebtnGravar{ PROP:XPos } = 0{ PROP:Width } - 90
      ?phDebugQueuebtnGravar{ PROP:YPos } = ?phDebugQueuebtnIgnorar{ PROP:YPos }
      ?phDebugQueuebtnFechar{ PROP:XPos } = 0{ PROP:Width } - 46
      ?phDebugQueuebtnFechar{ PROP:YPos } = ?phDebugQueuebtnIgnorar{ PROP:YPos }

    OF EVENT:Accepted
      CASE FIELD()
      OF ?phDebugQueuebtnIgnorar
        bignoredebugqueue = TRUE                            !Altera a baliza para ignorar próximas chamadas
        BREAK
      OF ?phDebugQueuebtnFechar
        BREAK
      OF ?phDebugQueuebtnGravar
        phExportQueue( pQ, pFile, pFormat, pFieldSeparator, pStringDelimiter, pRecNumbers )
      END
    OF EVENT:CloseWindow
    END
  END
  lLastX = 0{ PROP:XPos }                                   !Salva a posição da janela, para a próxima chamada
  lLastY = 0{ PROP:YPos }
  GET( pQ, lSavPointer )                                    !Reposiciona o ponteiro da queue

FormatQueue             ROUTINE
  ?phDebugQueueListBox1{ PROP:LineHeight } = 11                         !Ajusta a altura das linhas
  i# = 1                                                    !Começa pelo primeiro componente da queue
  LOOP                                                      !Inicia um loop contínuo
    pComponent &= WHAT( pQ, i# )                            !Determina o ponteiro para o componente da vez
    IF pComponent &= NULL THEN BREAK END                    !Determina se chegou ao final das colunas
    zstrColuna = WHO( pQ, i# )                              !Extrai o nome da coluna

    ?phDebugQueueListBox1{ PROPLIST:Header, i# } = zstrColuna
    ?phDebugQueueListBox1{ PROPLIST:Picture, i# } = '@s50'
    ?phDebugQueueListBox1{ PROPLIST:LeftOffset, i# } = 2
    ?phDebugQueueListBox1{ PROPLIST:HeaderLeftOffset, i# } = 2
    ?phDebugQueueLIstBox1{ PROPLIST:Underline, i# } = TRUE
    ?phDebugQueueListBox1{ PROPLIST:Resize, i# } = TRUE
    ?phDebugQueueListBox1{ PROPLIST:RightBorder + PROPLIST:Group, i# } = TRUE

    IF NOT ISSTRING( pComponent )                           !Se o campo não é uma string
      ?phDebugQueueListBox1{ PROPLIST:Right, i# } = TRUE
      ?phDebugQueueListBox1{ PROPLIST:HeaderRight, i# } = TRUE
      ?phDebugQueueListBox1{ PROPLIST:HeaderRightOffset, i# } = 2
    END

    i# += 1                                                 !Incrementa o contador de componentes
  END

! Isto distribui as colunas igualmente no listbox
  lColumnWidth = ?phDebugQueueListBox1{ PROP:Width } / i# - 1           !Calcula a largura de cada coluna, igualmente
  IF lColumnWidth < 10 THEN lColumnWidth = 10 .             !Não deixa que a largura seja muito pequena
  LOOP j# = 1 TO i#                                         !Percorre a queue novamente para acertar as larguras
    ?phDebugQueueListBox1{ PROPLIST:Width, j# } = lColumnWidth          !Aplica a largura calculada
  END

phExportQueue           PROCEDURE( QUEUE pQ, <STRING pFile>, BYTE pFormat=1, <STRING pFieldSeparator>, <STRING pStringDelimiter>, BYTE pRecNumbers=TRUE )
fASCIIph                FILE,DRIVER('ASCII','CLIP=ON'),CREATE
RECORD                    RECORD
l                           STRING( 65520 )
                          END
                        END

lExtensionIndex         LONG                                !Índice da extensão escolhida pelo usuário
zstrTmp                 CSTRING( 65520 )
zstrASCIIFileName       CSTRING( 161 )                      !Nome do arquivo ASCII a gerar (exclusivo desta procedure)
zstrFieldSeparator      STRING( ';' )                       !Separador de campos do arquivo CSV
zstrStringDelimiter     STRING( '"' )                       !Delimitador de strings
pComponent              ANY                                 !Ponteiro para a coluna da queue
zstrColuna              CSTRING( 50 )                       !Nome da coluna
lColunas                LONG,AUTO                           !Contador interno do número de colunas
qColunas                QUEUE                               !Queue de colunas da queue inspecionada
nome                      STRING( 50 )                      !Nome da coluna
tipo                      STRING( 1 )                       !Tipo de dado da coluna (string/numérica)
tamanho                   LONG                              !Comprimento máximo de dado encontrado
                        END
lSavPointer             LONG,AUTO                           !Posição original do ponteiro da queue

  CODE
  IF pFile THEN zstrASCIIFileName = pFile END               !Se foi passado um nome de arquivo

  IF NOT zstrASCIIFileName                                  !Se não existe um nome de arquivo informado
    IF NOT FILEDIALOGA( 'Informe o arquivo a criar', |      !Chama o diálogo padrão do Windows
                         zstrASCIIFileName, |
                        'Separado por Ponto e Vírgula;*.CSV|Arquivo Texto Posicional|*.TXT', |
                         FILE:Save + FILE:KeepDir + FILE:LongName, |
                         lExtensionIndex )
      RETURN FALSE
    END
  ELSE
    lExtensionIndex = pFormat                               !Carrega o formato desejado, vindo do parâmetro
  END

  fASCIIph{ PROP:Name } = zstrASCIIFileName
  IF EXISTS( zstrASCIIFileName )                            !Se o arquivo existe
    IF STATUS( fASCIIph ) <> 0                              !e está sendo utilizado,
      MESSAGE( 'O arquivo especificado está em uso', 'Gravar Queue', ICON:Hand ) !Mostra a mensagem ao usuário
      RETURN FALSE
    END
  END
  CREATE( fASCIIph )                                        !Cria o arquivo
  IF ERROR()                                                !Se houve qualquer erro na criação do arquivo
    MESSAGE( 'Erro na criação do arquivo ' & zstrASCIIFileName, 'Gravar Queue', ICON:Hand ) !Mostra a mensagem ao usuário
    RETURN FALSE
  END
  OPEN( fASCIIph )                                          !Abre o arquivo informado
  IF ERROR()                                                !Se houve qualquer erro na abertura do arquivo
    MESSAGE( 'Erro na abertura do arquivo ' & zstrASCIIFileName, 'Gravar Queue', ICON:Hand ) !Mostra a mensagem ao usuário
    RETURN FALSE
  END

  IF pFieldSeparator THEN zstrFieldSeparator = pFieldSeparator .
  IF pStringDelimiter THEN zstrStringDelimiter = pStringDelimiter .

  lSavPointer = POINTER( pQ )                               !Salva a posição do ponteiro da queue

! Determina quantas colunas existem na queue
  lColunas = 0
  k# = 1
  LOOP
    zstrColuna = WHO( pQ, k# )
    IF zstrColuna = '' THEN BREAK .
    qColunas.nome = zstrColuna                              !Salva o nome da coluna
    pComponent &= WHAT( pQ, k# )                            !Obtém um ponteiro para o dado da coluna
    IF ISSTRING( pComponent )
      qColunas.tipo = 'S'                                   !Coluna tipo string
    ELSE
      qColunas.tipo = 'N'                                   !Coluna tipo numérica
    END
    ADD( qColunas )                                         !Cria uma entrada na queue de colunas (se precisar gravar a queue)
    lColunas += 1                                           !Incrementa o número de colunas
    k# += 1
  END

  DO CheckTamanhos

  CASE lExtensionIndex                                      !Qual foi o tipo de extensão que o usuário selecionou?
  OF 1                                                      !Texto separado por vírgula
    IF pRecNumbers                                          !Se precisa incluir os números dos registros
      zstrTmp = 'REGISTRO;'                                 !Cria o nome da coluna
    END
    LOOP i# = 1 TO lColunas
      GET( qColunas, i# )
      zstrTmp = zstrTmp & CLIP( qColunas.nome )             !Cria o nome da coluna
      IF i# < lColunas                                      !Se ainda restam colunas a serem processadas
        zstrTmp = zstrTmp & zstrFieldSeparator              !Adiciona o separador ao final da linha
      END
    END
    fASCIIph.l = zstrTmp
    APPEND( fASCIIph )                                      !Grava a primeira linha (cabeçalho)
    LOOP i# = 1 TO RECORDS( pQ )                            !Percorre as linhas da queue
      zstrTmp = ''                                          !Limpa a linha de trabalho
      IF pRecNumbers                                        !Se precisa incluir os números dos registros
        zstrTmp = i# & ';'                                  !Numera a linha
      END
      GET( pQ, i# )
      LOOP j# = 1 TO lColunas                               !Percorre as colunas de cada linha
        GET( qColunas, j# )
        IF qColunas.tipo = 'S'
          zstrTmp = zstrTmp & zstrStringDelimiter & CLIP( WHAT( pQ, j# ) ) & zstrStringDelimiter
        ELSE
          zstrTmp = zstrTmp & WHAT( pQ, j# )
        END
        IF j# < lColunas                                    !Se ainda restam mais colunas para serem processadas
          zstrTmp = zstrTmp & zstrFieldSeparator
        END
      END
      fASCIIph.l = zstrTmp
      APPEND( fASCIIph )                                    !Inclui a linha no arquivo
    END

  OF 2                                                      !Texto posicional
    ult# = 1                                                !Primeira coluna utilizada
    zstrTmp = 'Colunas - início : fim'
    fASCIIph.l = zstrTmp
    APPEND( fASCIIph )
    LOOP i# = 1 TO lColunas
      GET( qColunas, i# )
      zstrTmp = CLIP( qColunas.nome ) & ' - ' & ult# & ' : ' & ( ult# + qColunas.tamanho - 1 )
      ult# += qColunas.tamanho                              !Calcula a última posição ocupada
      fASCIIph.l = zstrTmp
      APPEND( fASCIIph )
    END
    zstrTmp = '<13,10>Dados'
    fASCIIph.l = zstrTmp
    APPEND( fASCIIph )
    LOOP i# = 1 TO RECORDS( pQ )                            !Percorre as linhas da queue
      GET( pQ, i# )
      ult# = 1
      zstrTmp = ''
      LOOP j# = 1 TO lColunas                               !Percorre as colunas de cada linha
        GET( qColunas, j# )
        zstrTmp[ ult# : ult# + qColunas.tamanho ] = CLIP( WHAT( pQ, j# ) ) !Mapeia a linha
        ult# += qColunas.tamanho                            !Calcula a próxima coluna a ser utilizada
      END
      fASCIIph.l = zstrTmp                                  !Quando terminarem as colunas, transfere para o buffer da linha
      APPEND( fASCIIph )
    END
  END
  CLOSE( fASCIIph )                                         !Fecha o arquivo
  GET( pQ, lSavPointer )                                    !Reposiciona o ponteiro da queue

CheckTamanhos           ROUTINE
  LOOP i# = 1 TO RECORDS( pQ )                              !Percorre todas as LINHAS da queue
    GET( pQ, i# )
    LOOP j# = 1 TO lColunas                                 !Percorre todas as COLUNAS de cada linha
      pComponent &= WHAT( pQ, j# )
      GET( qColunas, j# )                                   !Acessa o registro da queue de colunas, referenta à coluna sendo lida
      IF ISSTRING( pComponent )                             !Se a coluna testada for do tipo string
        IF LEN( CLIP( pComponent ) ) > qColunas.tamanho
          qColunas.tamanho = LEN( CLIP( pComponent ) )
        END
      ELSE                                                  !A coluna não é do tipo string
        c# = INT( LOG10( pComponent ) ) + 1                 !Calcula o tamanho necessário em caracteres
        IF c# > qColunas.tamanho                            !Se o conteúdo da coluna é maior que o que está lá
          qColunas.tamanho = c#
        END
      END
      PUT( qColunas )                                       !Atualiza o registro da coluna em questão
    END
  END
