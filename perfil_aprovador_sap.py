import win32com.client  
import pandas as pd 
import time


##Conexão Python x SAP GUI
def conectar_sap():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session

## Carrega Planilha Excel 
def carregar_planilha():
    caminho = r"C:\Users\caio.vinicius\OneDrive - CantuStore\Área de Trabalho\SAP\SCRIPTS\PYTHON\Perfil_Aprovador_SAP\profile_approver.xlsm"
    aba = "PERFIL APROVADOR"
    try:
        df = pd.read_excel(caminho, sheet_name=aba)
        return df
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

##Criação Perfil Aprovador
def criar_perfil_aprovador(session, nome, sobrenome, usuario, id,sexo):
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "PA30"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtRP50G-PERNR").text = id
    session.findById("wnd[0]/usr/ctxtRP50G-PERNR").caretPosition = 6
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU").getAbsoluteRow(0).selected = True
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,0]").setFocus()
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITMENU:SAPMP50A:0310/tblSAPMP50ATC_MENU/txtGV_ITEXT[0,0]").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    session.findById("wnd[0]/usr/ctxtP0000-BEGDA").text = "01.01.2018"
    session.findById("wnd[0]/usr/ctxtPSPAR-WERKS").text = "BR01"
    session.findById("wnd[0]/usr/ctxtPSPAR-PERSG").text = "1"
    session.findById("wnd[0]/usr/ctxtPSPAR-PERSK").text = "BA"
    session.findById("wnd[0]/usr/ctxtPSPAR-PERSK").setFocus ()
    session.findById("wnd[0]/usr/ctxtPSPAR-PERSK").caretPosition = 2
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (11)
    session.findById("wnd[1]/tbar[0]/btn[12]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    if sexo.upper() == "F":
        session.findById("wnd[0]/usr/cmbQ0002-ANREX").key = "Sra."
        session.findById("wnd[0]/usr/radQ0002-GESC2").select()
        
    elif sexo.upper() == "M":
        session.findById("wnd[0]/usr/cmbQ0002-ANREX").key = "Sr."
        session.findById("wnd[0]/usr/radQ0002-GESC1").select()
    
    else:
        print("Preencha a coluna de sexo com F ou M apenas!")
        
        
    session.findById("wnd[0]/usr/txtP0002-NACHN").text = sobrenome
    session.findById("wnd[0]/usr/txtP0002-VORNA").text = nome
    session.findById("wnd[0]/usr/ctxtP0002-GBDAT").text = "01.01.1990"
    session.findById("wnd[0]/usr/ctxtP0002-NATIO").text = "BR"
    session.findById("wnd[0]/usr/txtT002T-SPTXT").setFocus()
    session.findById("wnd[0]/usr/txtT002T-SPTXT").caretPosition = 0
    session.findById("wnd[0]").sendVKey (11)
    session.findById("wnd[1]/tbar[0]/btn[12]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    session.findById("wnd[0]/usr/ctxtP0001-BTRTL").text = "0001"
    session.findById("wnd[0]/usr/ctxtP0001-GSBER").text = "0001"
    session.findById("wnd[0]/usr/ctxtP0001-ABKRS").text = "01"
    session.findById("wnd[0]/usr/ctxtP0001-ABKRS").setFocus()
    session.findById("wnd[0]/usr/ctxtP0001-ABKRS").caretPosition = 2
    session.findById("wnd[0]/usr/btnMOREPLAN").press()
    session.findById("wnd[1]/usr/radCROSS_ORGEH").select()
    session.findById("wnd[1]/usr/ctxtASS_ORGEH").setFocus()
    session.findById("wnd[1]/usr/ctxtASS_ORGEH").caretPosition = 0
    session.findById("wnd[1]").sendVKey (2)
    session.findById("wnd[2]").close()
    session.findById("wnd[1]/usr/ctxtASS_ORGEH").text = "50000000"
    session.findById("wnd[1]/usr/ctxtASS_ORGEH").setFocus()
    session.findById("wnd[1]/usr/ctxtASS_ORGEH").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]").sendVKey (11)
    session.findById("wnd[0]/usr/txtP0006-STRAS").text = "RUA A"
    session.findById("wnd[0]/usr/txtP0006-HSNMR").text = "10"
    session.findById("wnd[0]/usr/ctxtP0006-ORT01").text = "ITAJAI"
    session.findById("wnd[0]/usr/txtP0006-PSTLZ").text = "88316-001"
    session.findById("wnd[0]/usr/ctxtP0006-STATE").text = "SC"
    session.findById("wnd[0]/usr/ctxtP0006-STATE").SetFocus()
    session.findById("wnd[0]/usr/ctxtP0006-STATE").caretPosition = 2
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (11)
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (11)
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-CHOIC").Text = "105"
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-SUBTY").Text = "0001"
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-SUBTY").SetFocus()
    session.findById("wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350/ctxtRP50G-SUBTY").caretPosition = 4
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    session.findById("wnd[0]/usr/txtP0105-USRID").text = usuario
    session.findById("wnd[0]/usr/txtP0105-USRID").SetFocus()
    session.findById("wnd[0]/usr/txtP0105-USRID").caretPosition = 7
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").sendVKey (11)
    
    ##Rodar o programa gerador do perfil na SE38
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "SE38"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "/SHCM/RH_SYNC_BUPA_EMPL_SINGLE"
    session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = 30
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtPNPPERNR-LOW").text = id
    session.findById("wnd[0]/usr/ctxtPNPPERNR-LOW").SetFocus()
    session.findById("wnd[0]/usr/ctxtPNPPERNR-LOW").caretPosition = 8
    session.findById("wnd[0]").sendVKey (8)
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

def executar_processamento():

    df = carregar_planilha()  
    if df is None:
        return  

    session = conectar_sap()

    for index, row in df.iterrows():
        nome = str(row[0]).strip()
        sobrenome = str(row[1]).strip()
        usuario = str(row[2]).strip()
        id = str(row[3]).strip()
        sexo =str(row[4]).strip()

        if all([nome, sobrenome, usuario, id, sexo]):
            criar_perfil_aprovador(session,nome, sobrenome, usuario, id, sexo)

if __name__ == "__main__":
 ### Menu de Validação
    while True:
        print("=======================Menu Reset Senhas==========================")
        print("===============Selecione uma das opções abaixo====================")
        print("===============[1] Criar Perfil de Aprovador======================")
        print("===============[2] Sair========================================== ")
        print("==================================================================")
        menu_option = input("=> ")

        if menu_option  == "1":
            print(f"Opção [{menu_option}] selecionada com sucesso!\n")
            ## Executa a função de criar perfil de aprovador      
            executar_processamento()
            
        elif menu_option == "2":
             print(f"Opção [{menu_option}] selecionada com sucesso!\n")
             print("Saindo da aplicação...\n")
             break
            
        else:
            print("Selecione apenas uma das opções [1] ou [2].\n")
