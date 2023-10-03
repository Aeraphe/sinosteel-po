Attribute VB_Name = "action_bootstrap"


'namespace=vba-files\Actions


Const USER_CONF_START_ROW = 6

Function run()

    On Error GoTo ErrorHanler
   Call Xhelper.waitMs(1000)
   On Error GoTo ErrorHanler
   loda_app_configs
   Exit Function
   
ErrorHanler:
 Call Alert.Show("Favor Carregar as Configurações!!", "Inicializando", 2000)
End Function


Public Function loda_app_configs()
  
  Dim row As Long

  core_env.read_config_file
  
  row = USER_CONF_START_ROW

  For Each varKey In core_env.props.Keys()
    user_app_config_sheet.Cells(row, "B").Value = core_env.props(varKey)

    row = row + 1
  Next 's
  
    Call Alert.Show("Aplicativo Pronto para Usuo!!", "Configurações Carregadas", 2000)
    
    ThisWorkbook.Save
    
    
End Function
