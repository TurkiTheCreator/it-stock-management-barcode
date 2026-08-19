Attribute VB_Name = "Module_BarcodeScan"
'==========================================================================
' Module_BarcodeScan
' Prototype de macro pour l'application de gestion de stock IT par
' codes-barres. A importer dans Excel via l'editeur VBA (Alt+F11 >
' Fichier > Importer un fichier > Module_BarcodeScan.bas).
'
' Principe :
'   1. L'utilisateur scanne un code-barres (le lecteur agit comme un
'      clavier et tape le code suivi d'un retour a la ligne).
'   2. La macro recherche le numero de serie correspondant.
'   3. Si le numero de serie existe deja dans la feuille Stock_Inventory,
'      l'enregistrement est bloque et une alerte s'affiche (doublon).
'   4. Sinon, une nouvelle ligne est ajoutee et l'action est journalisee
'      dans la feuille Scan_Log.
'==========================================================================

Option Explicit

Private Const SHEET_STOCK As String = "Stock_Inventory"
Private Const SHEET_LOG As String = "Scan_Log"
Private Const STOCK_HEADER_ROW As Long = 4
Private Const STOCK_SERIAL_COL As Long = 2   ' Colonne B = Numero de serie

' Point d'entree principal : a lier a un bouton ou a un raccourci clavier
Sub ScanBarcode()
    Dim barcodeInput As String
    Dim serialInput As String
    Dim wsStock As Worksheet
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim isDuplicate As Boolean

    Set wsStock = ThisWorkbook.Sheets(SHEET_STOCK)
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOG)

    barcodeInput = InputBox("Scanner ou saisir le code-barres :", "Scan - Entree stock")
    If Trim(barcodeInput) = "" Then Exit Sub

    serialInput = InputBox("Numero de serie associe :", "Scan - Entree stock")
    If Trim(serialInput) = "" Then Exit Sub

    isDuplicate = SerialExists(wsStock, serialInput)

    If isDuplicate Then
        MsgBox "Doublon detecte !" & vbCrLf & _
               "Le numero de serie '" & serialInput & "' existe deja dans l'inventaire." & vbCrLf & _
               "Enregistrement bloque.", vbCritical, "Alerte doublon"
        LogScan wsLog, barcodeInput, serialInput, "Verification", "Bloque - doublon detecte"
        Exit Sub
    End If

    lastRow = wsStock.Cells(wsStock.Rows.Count, STOCK_SERIAL_COL).End(xlUp).Row + 1

    wsStock.Cells(lastRow, 1).Value = barcodeInput          ' Barcode ID
    wsStock.Cells(lastRow, 2).Value = serialInput            ' Numero de serie
    wsStock.Cells(lastRow, 6).Value = Format(Now, "dd/mm/yyyy") ' Date d'entree
    wsStock.Cells(lastRow, 8).Value = "En stock"              ' Statut

    ' Recopie la formule de detection de doublon sur la nouvelle ligne
    wsStock.Cells(lastRow, 9).Formula = _
        "=IF(COUNTIF($B$" & (STOCK_HEADER_ROW + 1) & ":$B" & lastRow & ",B" & lastRow & ")>1,""DOUBLON"",""OK"")"

    LogScan wsLog, barcodeInput, serialInput, "Entree stock", "Enregistre avec succes"

    MsgBox "Materiel enregistre avec succes.", vbInformation, "Scan reussi"
End Sub

' Recherche si un numero de serie existe deja dans la colonne des numeros de serie
Private Function SerialExists(ws As Worksheet, serial As String) As Boolean
    Dim lastRow As Long
    Dim found As Range

    lastRow = ws.Cells(ws.Rows.Count, STOCK_SERIAL_COL).End(xlUp).Row
    If lastRow < STOCK_HEADER_ROW + 1 Then
        SerialExists = False
        Exit Function
    End If

    Set found = ws.Range(ws.Cells(STOCK_HEADER_ROW + 1, STOCK_SERIAL_COL), _
                          ws.Cells(lastRow, STOCK_SERIAL_COL)).Find( _
                          What:=serial, LookAt:=xlWhole, MatchCase:=False)

    SerialExists = Not (found Is Nothing)
End Function

' Ajoute une ligne au journal des scans (feuille Scan_Log)
Private Sub LogScan(ws As Worksheet, barcode As String, serial As String, action As String, result As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(lastRow, 1).Value = Format(Now, "dd/mm/yyyy hh:mm")
    ws.Cells(lastRow, 2).Value = barcode
    ws.Cells(lastRow, 3).Value = serial
    ws.Cells(lastRow, 4).Value = action
    ws.Cells(lastRow, 5).Value = result
End Sub
