# -*- coding: utf-8 -*-
"""
Created on Tue Aug 30 15:00:17 2022

@author: Beto Schneider
"""


class SAP():
    def __init__(self):
        self.path = r'C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe'
        self.app = '01 PEP - SAP S/4HANA Produção'
        self.erro = None
        
        #bibliotecas
        import win32com.client
        import sys
        import subprocess
        import time
        
        try:
            subprocess.Popen(self.path)
            time.sleep(20) #aguarda 20 segundos para que o programa seja aberto pro completo
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return
            time.sleep(3)
            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            time.sleep(3)
            #faz login na aplicação selecionada
            connection = application.OpenConnection(self.app, True)
            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return 
            #descomentar linha abaixo caso seja exigida autenticação no SAP
            #time.sleep(20)
            self.session = connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return
        except:
            self.erro = sys.exc_info()
    
    #metodo para extrair dados para o Excel
    def extrair(self, transacao, pasta, cado_dtini=None, cado_dtfim=None):
        '''Transações: CN43N, CN47N, YSRELCONT, CADO e KS13'''
        import sys
        import time
        session = self.session
        transacao = transacao.upper()
        try:
            if transacao == 'CN43N':
                time.sleep(3)
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/tbar[0]/okcd").text = '/n' + transacao
                time.sleep(5) # somente pra verificação visual
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
                try:
                    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "000000000001"
                    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").caretPosition = 12
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except:
                    pass
                session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = "PT-200*99*"
                session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/CENTRAL"
                session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus()
                session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 8
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").setCurrentCell(10,"POST1")
                session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = transacao + '.XLSX'
                #session.findById("wnd[1]/tbar[0]/btn[0]").press() # botão gerar
                session.findById("wnd[1]/tbar[0]/btn[11]").press() # botão substituir
                # ir para tela inicial com comando /n
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
                
            if transacao == 'CN47N':
                time.sleep(3)
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/tbar[0]/okcd").text = '/n' + transacao
                #time.sleep(2) # somente pra verificação visual
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
                try:
                    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = "000000000001"
                    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").caretPosition = 12
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except:
                    pass
                session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = "PT-200*"
                session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/CENTRAL"
                session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus()
                session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 8
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").setCurrentCell(8,"LTXA1")
                session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = transacao + '.XLSX'
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
                #session.findById("wnd[1]/tbar[0]/btn[0]").press() # botão gerar
                session.findById("wnd[1]/tbar[0]/btn[11]").press() # botão substituir
                # ir para tela inicial com comando /n
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
            
            if transacao == 'KS13':
                time.sleep(3)
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/tbar[0]/okcd").text = '/n' + transacao
                #time.sleep(2) # somente pra verificação visual
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/subKOSTL_SELECTION:SAPLKMS1:0100/radKMAS_D-KZKOSTLALL").setFocus()
                session.findById("wnd[0]/usr/subKOSTL_SELECTION:SAPLKMS1:0100/radKMAS_D-KZKOSTLALL").select()
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]").sendVKey(33)
                session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").currentCellRow = 56 #linha do layout /ANALITICO
                session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").firstVisibleRow = 50
                session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "56" #linha do layout /ANALITICO
                session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell() #remover parenteses
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(8,"KTEXT")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta #r"C:\Users\DJK8\PETROBRAS\Central - Histórico de extrações\20230328"
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = transacao + '.XLSX' #"KS13.XLSX"
                session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
                session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 68
                session.findById("wnd[1]").sendVKey(4)
                session.findById("wnd[2]/tbar[0]/btn[11]").press()
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                
            if transacao == 'YSRELCONT':
                time.sleep(3)
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/tbar[0]/okcd").text = '/n' + transacao
                #time.sleep(2) # somente pra verificação visual
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
                session.findById("wnd[0]/usr/ctxtSC_BSART-LOW").text = "ZCVR"
                session.findById("wnd[0]/usr/ctxtSC_EKORG-LOW").text = "0000"
                session.findById("wnd[0]/usr/ctxtSC_EKORG-HIGH").text = "9999"
                session.findById("wnd[0]/usr/ctxtSD_KDATE-LOW").text = "01.01.2020"
                session.findById("wnd[0]/usr/ctxtSD_KDATE-HIGH").text = "01.01.2099"
                session.findById("wnd[0]/usr/ctxtP_VARIA").text = "/CENTRAL2"
                session.findById("wnd[0]/usr/ctxtP_VARIA").setFocus()
                session.findById("wnd[0]/usr/ctxtP_VARIA").caretPosition = 9
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]").sendVKey(33)
                session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 177
                session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 168
                session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "177"
                session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(11,"TXCAB")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "11" #11
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = transacao + '.XLSX'
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
                #session.findById("wnd[1]/tbar[0]/btn[0]").press() # botão gerar
                session.findById("wnd[1]/tbar[0]/btn[11]").press() # botão substituir
                # ir para tela inicial com comando /n
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
            
            if transacao == 'CADO':
                time.sleep(3)
                session.findById("wnd[0]").maximize()
                session.findById("wnd[0]/tbar[0]/okcd").text = '/n' + transacao
                #time.sleep(2) # somente pra verificação visual
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
                session.findById("wnd[0]/tbar[1]/btn[27]").press()
                if cado_dtini != None and cado_dtfim != None:
                    session.findById("wnd[0]/usr/ctxtSO_DATUM-LOW").text = cado_dtini
                    session.findById("wnd[0]/usr/ctxtSO_DATUM-HIGH").text = cado_dtfim
                else:
                    session.findById("wnd[0]/usr/radMONAT").select()
                session.findById("wnd[0]/usr/ctxtSO_SKOSL-LOW").text = "SP00ADCB17"
                session.findById("wnd[0]/usr/ctxtSO_SKOSL-HIGH").text = "SPBIH1AB17"
                session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CENTRAL2"
                session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
                session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 9
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(12,"RNPLNR")
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = transacao + '.XLSX'
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
                #session.findById("wnd[1]/tbar[0]/btn[0]").press() # botão gerar
                session.findById("wnd[1]/tbar[0]/btn[11]").press() # botão substituir
                # ir para tela inicial com comando /n
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
        except:
            self.erro = [transacao, sys.exc_info()]

    def logoff(self):
        import sys
        import time
        session = self.session
        try:
            time.sleep(5)
            session.findById("wnd[0]").maximize()
            #saída da aplicação sistema sem pedido de confirmação 
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            session.findById("wnd[0]").sendVKey(0)
            
            global connection
            global application
            global SapGuiAuto
            
            connection = None
            application = None
            SapGuiAuto = None
            self.session = None
            
        except:
            self.erro = sys.exc_info()


if __name__ == '__main__':
    
    #instancia objeto sap
    sap = SAP()
    
    #local de salvamento dos arquivos extraidos
    pasta = r'C:\Users\DJK8\Downloads'
    
    #metodo extrair para cada transação
    sap.extrair('CN43N', pasta)
    #sap.extrair('CN47N', pasta)
    #sap.extrair('KS13', pasta)
    #sap.extrair('YSRELCONT', pasta)
    #sap.extrair('CADO', pasta)
    #sap.extrair('CADO', pasta, cado_dtini='01.08.2022', cado_dtfim='15.08.2022')
    sap.logoff()
    
    if sap.erro != None:
        print(sap.erro)