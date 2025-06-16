import win32com.client  



### Conexão Python x SAP GUI
def sap_conection():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session

### Reset de Senha Unitária
def password_reset(session, sap_user, sap_password):
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "su01"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = sap_user
    session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").caretPosition = 10
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[18]").press()
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO").select()
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD").text = sap_password
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").text = sap_password
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").caretPosition = 10    
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/tbar[0]/btn[6]").press()
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/tbar[0]/btn[9]").press()
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/tbar[0]/btn[9]").press()
    session.findById("wnd[0]/tbar[1]/btn[20]").press()
    session.findById("wnd[1]/usr/subPOPUP:SAPLSUID_MAINTENANCE:2101/cntlG_CUA_SYSTEMS_CONTAINER1/shellcont/shell").setCurrentCell (-1,"")
    session.findById("wnd[1]/usr/subPOPUP:SAPLSUID_MAINTENANCE:2101/cntlG_CUA_SYSTEMS_CONTAINER1/shellcont/shell").selectAll()
    session.findById("wnd[1]/usr/subPOPUP:SAPLSUID_MAINTENANCE:2101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD").text = sap_password
    session.findById("wnd[1]/usr/subPOPUP:SAPLSUID_MAINTENANCE:2101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").text = sap_password
    session.findById("wnd[1]/usr/subPOPUP:SAPLSUID_MAINTENANCE:2101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").setFocus()
    session.findById("wnd[1]/usr/subPOPUP:SAPLSUID_MAINTENANCE:2101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    
    
### Reset de Senha em Massa
def password_reset_mass(session, sap_user, sap_password):
    pass
    
### Menu de Validação
while True:
    print("====================Menu Reset Senhas====================")
    print("=============Selecione uma das opções abaixo=============")
    print("=============[1] Reset Senha Unitária [SU01]=============")
    print("=============[2] Reset Senha em Massa [SU10]=============")
    print("=========================================================")
    menu_option = input("=> ")

    if menu_option in ["1", "2"]:
        print(f"Opção [{menu_option}] selecionada com sucesso!\n")
        ### Inserção de Dados do Usuário
        sap_user = str(input("Digite o nome do usuário: \n"))
        sap_password = str(input("Digite a senha: \n"))
    
         ## Executa a funçção de verificação da instãncia do SAP
        session = sap_conection()
        
        password_reset(session, sap_user, sap_password)
        
    else:
        print("Selecione apenas uma das opções [1] ou [2].\n")
        break

### Executa o Script de acordo com o valor selecionado no Menu de Validação
if menu_option == "1":
    password_reset(session, sap_user, sap_password)
    print(f"Usuário {sap_user} resetada com sucesso!!!!")
elif menu_option == "2":
    password_reset_mass(session, sap_user, sap_password)
    print(f"Usuários {sap_user} resetados com sucesso!!!!")
    







    







    

