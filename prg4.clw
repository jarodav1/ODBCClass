prg4                    PROGRAM                             ! Vínculo de controles com variáveis
                                                            ! Vantagem: alto dinamismo
                        MAP                                 ! Desvantagem: difícil prever todas as situações, você pode ficar louco
                          fPrepare
                          SelectPessoa
                        END
                        INCLUDE( 'keycodes.clw' )

wdw                     WINDOW( 'Vínculo entre controles e variáveis' ), CENTER, resize, GRAY
                          PROMPT( '&ID:' ), AT( 5, 5 ), TRN ; ENTRY( @N10. ), USE( ?id ), AT( 50, 5 )
                          PROMPT( '&Nome: ' ), AT( 5, 25 ), TRN ; ENTRY( @S50 ), USE( ?nome ), AT( 50, 25 )

                          BUTTON( '<< &Anterior' ), USE( ?btnAnterior ), AT( 230, 5, 40, 12 ), KEY( PgUpKey ), DISABLE
                          BUTTON( '&Próximo >' ), USE( ?btnProximo ), AT( 275, 5, 40, 12 ), KEY( pgDnKey )
                          BUTTON( '&Gravar' ), USE( ?btnGravar ), AT( 275, 25, 40, 12 ), KEY( CtrlEnter )
                        END

f                       FILE, DRIVER( 'TopSpeed' ), NAME( 'prg4.tps' ), CREATE
f_pk                      KEY( id ), PRIMARY
RECORD                    RECORD
id                          LONG
nome                        STRING( 50 )
                          END
                        END

wdwClassVars            QUEUE, TYPE
value                     ANY
feq                       LONG
                        END

wdwClass                CLASS, TYPE
Vars                      &wdwClassVars
AddVar                    PROCEDURE( *? pVar, LONG pFeq )
Construct                 PROCEDURE
Destruct                  PROCEDURE
                        END

TW                      wdwClass
lPessoa                 LONG( 1 )

  CODE
  fPrepare
  OPEN( wdw )

  TW.AddVar( f.id, ?id )
  TW.AddVar( f.nome, ?nome )

  ACCEPT
    DISPLAY()

    CASE EVENT()
    OF EVENT:Accepted
      CASE FIELD()
      OF ?btnAnterior
        lPessoa -= 1
        SelectPessoa
        IF lPessoa = 1 THEN DISABLE( ? ) END
        ENABLE( ?btnProximo )

      OF ?btnProximo
        lPessoa += 1
        SelectPessoa
        IF lPessoa = RECORDS( f ) THEN disable( ? ) END
        ENABLE( ?btnAnterior )

      OF ?btnGravar
        PUT( f )
        IF NOT ERRORCODE()
          MESSAGE( 'Registro atualizado com sucesso' )
        END

      ELSE                                                  ! Nenhum botão previamente conhecido
        TW.Vars.feq = FIELD()
        GET( TW.Vars, TW.Vars.feq )
        IF NOT ERRORCODE()                                  ! Encontrou o controle que gerou o evento
          MESSAGE( 'ID:<9>' & f.id & '<13,10>Nome:<9>' & CLIP( f.Nome ), 'Registro', ICON:Exclamation )
        END

      END
    END

  END

wdwClass.AddVar         PROCEDURE( *? pVar, LONG pFeq )
  CODE
  CLEAR( SELF.Vars )                                        ! Precisa dar um clear antes de adicionar
  SELF.Vars.value &= pVar
  SELF.Vars.feq = pFeq
  ADD( SELF.Vars )
  pFeq{ PROP:Use } = pVar

wdwClass.Construct      PROCEDURE
  CODE
  SELF.Vars &= NEW wdwClassVars

wdwClass.Destruct       PROCEDURE
  CODE
  LOOP i# = 1 TO RECORDS( SELF.Vars )
    GET( SELF.Vars, i# )
    SELF.Vars.value &= NULL
  END
  DISPOSE( SELF.Vars )

fPrepare                PROCEDURE
  CODE
  IF NOT EXISTS( f{ PROP:Name } )
    CREATE( f )
  END
  OPEN( f, 42h )
  IF NOT RECORDS( f )
    LOOP i# = 1 TO 10
      f.id = i#
      f.nome = 'PESSOA ' & i#
      ADD( f )
    END
  END
  SelectPessoa

SelectPessoa            PROCEDURE
  CODE
  f.id = lPessoa
  GET( f, f.f_pk )
