Attribute VB_Name = "SmoothyInventoryApp"
Option Explicit

' =====================================================================================
' SmoothyInventoryApp - Multi-business (Video Games + Pokemon + Cologne) - ALL FEATURES
' CLEAN / COMPILE-READY SINGLE MODULE
'
' IMPORTANT FIX (what you asked for):
' - Console / Product dropdown is now properly linked everywhere:
'     Settings -> named ranges (Field_console___product_*List)
'     Add Stock -> binds Console / Product dropdown based on selected Business
'     Batch Entry -> binds per-row Console / Product dropdown based on Business column
'     Inventory sheet -> applies Console / Product validation per row using that row's Business
'
' Notes:
' - Excel Data Validation does NOT auto-switch list sources when the Business cell changes
'   unless you use sheet event code. In THIS .bas we provide:
'     RebindAddStockDropdowns
'     RebindBatchDropdowns
'     RebindInventoryDropdowns
'   Run them after changing business values / after paste/import.
' =====================================================================================

' =========================
' Types
' =========================
Public Type InventoryCols
    sku As Long
    Console As Long            ' Console / Product
    GameName As Long
    DateAdded As Long
    Country As Long
    Quality As Long
    Quantity As Long
    paid As Long
    FromCol As Long
    Seller As Long
    ListedStatus As Long
    ListedPrice As Long
    Tested As Long
    Bundle As Long
    Stored As Long
    Notes As Long
    Business As Long
End Type

' =========================
' BRAND / TEXT
' =========================
Public Const BRAND_TITLE As String = "Smoothy Game Store Dashboard"
Public Const BRAND_SUB As String = "Pokemon. Video Games and Fun!"
Public Const MSG_TITLE As String = "Smoothy Inventory"

' =========================
' Layout
' =========================
Private Const NORMAL_ROW1_HEIGHT As Double = 18
Private Const NORMAL_ROW2_HEIGHT As Double = 15

' =========================
' Theme / Colors
' =========================
Public THEME_FONT As String
Public CLR_PRIMARY As Long, CLR_PRIMARY_DARK As Long, CLR_SECONDARY As Long, CLR_BG As Long
Public CLR_PANEL As Long, CLR_PANEL_ALT As Long, CLR_BORDER As Long
Public CLR_ACCENT As Long, CLR_ACCENT_ALT As Long
Public CLR_SUCCESS As Long, CLR_ERROR As Long, CLR_NEUTRAL As Long
Public CLR_ALTROW As Long, CLR_INPUT As Long, CLR_HEADER As Long
Public CLR_BUTTON_BG As Long, CLR_BUTTON_BG_HOVER As Long, CLR_BUTTON_TEXT As Long, CLR_BUTTON_BORDER As Long
Public CLR_HISTORY_HEADER_BG As Long, CLR_HISTORY_HEADER_TEXT As Long

Public g_IsDarkMode As Boolean
Private g_ThemeApplied As Boolean

' =========================
' Businesses
' =========================
Private Const BIZ_VG As String = "Video Games"
Private Const BIZ_PK As String = "Pokemon"
Private Const BIZ_CL As String = "Cologne"
Private Const BIZ_ALL As String = "All"

' SKU prefixes
Private Const SKU_PREFIX_VG As String = "SKUVG"
Private Const SKU_PREFIX_PK As String = "SKUPK"
Private Const SKU_PREFIX_CL As String = "SKUCL"

' Config keys
Private Const CFG_SKU_VG As String = "LastSKUIndex_VG"
Private Const CFG_SKU_PK As String = "LastSKUIndex_PK"
Private Const CFG_SKU_CL As String = "LastSKUIndex_CL"
Private Const CFG_ACTIVE_BIZ As String = "ActiveBusiness"

' =====================================================================================
' THEME
' =====================================================================================
Private Sub EnsureThemeReady()
    If g_ThemeApplied And THEME_FONT <> "" And CLR_ALTROW <> 0 And CLR_PRIMARY <> 0 Then Exit Sub
    SetThemeLight
End Sub

Public Sub SetThemeLight()
    THEME_FONT = "Segoe UI"
    CLR_PRIMARY = RGB(18, 67, 129)
    CLR_PRIMARY_DARK = RGB(10, 46, 92)
    CLR_SECONDARY = vbWhite
    CLR_BG = RGB(246, 249, 253)
    CLR_PANEL = RGB(230, 238, 248)
    CLR_PANEL_ALT = RGB(218, 230, 244)
    CLR_BORDER = RGB(158, 175, 194)
    CLR_ACCENT = RGB(0, 120, 215)
    CLR_ACCENT_ALT = RGB(0, 150, 255)
    CLR_SUCCESS = RGB(80, 165, 80)
    CLR_ERROR = RGB(215, 70, 85)
    CLR_NEUTRAL = RGB(155, 160, 170)
    CLR_ALTROW = RGB(234, 241, 250)
    CLR_INPUT = RGB(255, 253, 210)
    CLR_HEADER = RGB(205, 220, 235)
    CLR_BUTTON_BG = RGB(32, 104, 180)
    CLR_BUTTON_BG_HOVER = RGB(50, 130, 205)
    CLR_BUTTON_TEXT = vbWhite
    CLR_BUTTON_BORDER = RGB(18, 68, 125)
    CLR_HISTORY_HEADER_BG = CLR_PRIMARY_DARK
    CLR_HISTORY_HEADER_TEXT = CLR_SECONDARY
    g_IsDarkMode = False
    g_ThemeApplied = True
End Sub

Public Sub SetThemeDark()
    THEME_FONT = "Segoe UI"
    CLR_PRIMARY = RGB(52, 78, 115)
    CLR_PRIMARY_DARK = RGB(32, 50, 74)
    CLR_SECONDARY = RGB(250, 250, 252)
    CLR_BG = RGB(32, 36, 46)
    CLR_PANEL = RGB(54, 62, 78)
    CLR_PANEL_ALT = RGB(64, 72, 90)
    CLR_BORDER = RGB(95, 110, 130)
    CLR_ACCENT = RGB(0, 160, 255)
    CLR_ACCENT_ALT = RGB(90, 200, 255)
    CLR_SUCCESS = RGB(90, 185, 90)
    CLR_ERROR = RGB(220, 90, 110)
    CLR_NEUTRAL = RGB(185, 190, 200)
    CLR_ALTROW = RGB(50, 58, 72)
    CLR_INPUT = RGB(110, 100, 30)
    CLR_HEADER = RGB(85, 100, 120)
    CLR_BUTTON_BG = RGB(60, 110, 185)
    CLR_BUTTON_BG_HOVER = RGB(85, 140, 210)
    CLR_BUTTON_TEXT = RGB(255, 255, 255)
    CLR_BUTTON_BORDER = RGB(30, 70, 120)
    CLR_HISTORY_HEADER_BG = CLR_PRIMARY_DARK
    CLR_HISTORY_HEADER_TEXT = CLR_SECONDARY
    g_IsDarkMode = True
    g_ThemeApplied = True
End Sub

Public Sub ToggleDarkMode()
    If Not g_ThemeApplied Or THEME_FONT = "" Then SetThemeLight
    If g_IsDarkMode Then SetThemeLight Else SetThemeDark
    ReapplyThemeAll
    RepairInventoryColors
    MsgBox "Theme: " & IIf(g_IsDarkMode, "Dark", "Light"), vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' UTILITIES
' =====================================================================================
Private Function Norm(ByVal s As String) As String: Norm = LCase$(Trim$(s)): End Function

Private Sub NormalizeTopRows(ws As Worksheet)
    On Error Resume Next
    ws.Rows(1).RowHeight = NORMAL_ROW1_HEIGHT
    ws.Rows(2).RowHeight = NORMAL_ROW2_HEIGHT
End Sub

Private Sub ApplySheetChrome(ws As Worksheet, Optional clearExisting As Boolean = False)
    If clearExisting Then ws.Cells.Clear
    EnsureThemeReady
    ws.Cells.Font.Name = THEME_FONT
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = CLR_BG
    NormalizeTopRows ws
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add
        EnsureSheet.Name = sheetName
    End If
End Function

Private Function MatchCol(ws As Worksheet, ByVal headerText As Variant) As Long
    Dim m As Variant
    On Error Resume Next
    m = Application.Match(CStr(headerText), ws.Rows(1), 0)
    On Error GoTo 0
    If Not IsError(m) Then MatchCol = CLng(m)
End Function

Private Function MatchConsoleColAny(ws As Worksheet) As Long
    Dim c As Long
    c = MatchCol(ws, "Console / Product")
    If c = 0 Then c = MatchCol(ws, "Console")
    MatchConsoleColAny = c
End Function

Private Function TableLastRow(ws As Worksheet, firstColLetter As String) As Long
    TableLastRow = ws.Cells(ws.Rows.Count, firstColLetter).End(xlUp).Row
End Function

' =========================
' SafeRun wrapper
' =========================
Public Sub SafeRun(ByVal taskName As String, ByVal workerProc As String)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Run workerProc
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Error in " & taskName & ":" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, MSG_TITLE
    Resume CleanExit
End Sub

' =====================================================================================
' CONFIG
' =====================================================================================
Public Sub SetupConfig()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Config")
    If ws.Range("A1").value = "" Then
        ws.Range("A1").value = "Setting": ws.Range("B1").value = "Value"
        ws.Range("A2").value = "LowStockThreshold": ws.Range("B2").value = 5
        ws.Range("A3").value = "DefaultListedStatus": ws.Range("B3").value = "Listed On eBay"
    End If
    If GetConfigValue(CFG_SKU_VG, "") = "" Then SaveConfigValue CFG_SKU_VG, 0
    If GetConfigValue(CFG_SKU_PK, "") = "" Then SaveConfigValue CFG_SKU_PK, 0
    If GetConfigValue(CFG_SKU_CL, "") = "" Then SaveConfigValue CFG_SKU_CL, 0
    If GetConfigValue(CFG_ACTIVE_BIZ, "") = "" Then SaveConfigValue CFG_ACTIVE_BIZ, BIZ_ALL

    ApplySheetChrome ws, False
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetConfigValue(ByVal key As String, Optional defaultValue As Variant) As Variant
    Dim ws As Worksheet, f As Range
    Set ws = EnsureSheet("Config")
    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)
    If f Is Nothing Then GetConfigValue = defaultValue Else GetConfigValue = f.Offset(0, 1).value
End Function

Private Sub SaveConfigValue(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, f As Range
    Set ws = EnsureSheet("Config")
    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole)
    If f Is Nothing Then
        Set f = ws.Cells(ws.Rows.Count, "A").End(xlUp).Offset(1, 0)
        f.value = key
    End If
    f.Offset(0, 1).value = value
End Sub

' =====================================================================================
' BUSINESS HELPERS
' =====================================================================================
Private Function NormalizeBusiness(ByVal s As String) As String
    Dim t As String: t = Trim$(s)
    If LCase$(t) = LCase$(BIZ_VG) Then NormalizeBusiness = BIZ_VG: Exit Function
    If LCase$(t) = LCase$(BIZ_PK) Then NormalizeBusiness = BIZ_PK: Exit Function
    If LCase$(t) = LCase$(BIZ_CL) Then NormalizeBusiness = BIZ_CL: Exit Function
    If LCase$(t) = LCase$(BIZ_ALL) Then NormalizeBusiness = BIZ_ALL: Exit Function

    If LCase$(left$(t, Len(SKU_PREFIX_VG))) = LCase$(SKU_PREFIX_VG) Then NormalizeBusiness = BIZ_VG: Exit Function
    If LCase$(left$(t, Len(SKU_PREFIX_PK))) = LCase$(SKU_PREFIX_PK) Then NormalizeBusiness = BIZ_PK: Exit Function
    If LCase$(left$(t, Len(SKU_PREFIX_CL))) = LCase$(SKU_PREFIX_CL) Then NormalizeBusiness = BIZ_CL: Exit Function

    NormalizeBusiness = BIZ_VG
End Function

Private Function BusinessMatchesFilter(ByVal rowBiz As String, ByVal filterBiz As String) As Boolean
    filterBiz = NormalizeBusiness(filterBiz)
    If filterBiz = BIZ_ALL Then
        BusinessMatchesFilter = True
    Else
        BusinessMatchesFilter = (NormalizeBusiness(rowBiz) = filterBiz)
    End If
End Function

Private Function BusinessToPrefix(ByVal biz As String) As String
    Select Case NormalizeBusiness(biz)
        Case BIZ_PK: BusinessToPrefix = SKU_PREFIX_PK
        Case BIZ_CL: BusinessToPrefix = SKU_PREFIX_CL
        Case Else:   BusinessToPrefix = SKU_PREFIX_VG
    End Select
End Function

Private Function PrefixToBusiness(ByVal sku As String) As String
    sku = Trim$(sku)
    If left$(sku, Len(SKU_PREFIX_PK)) = SKU_PREFIX_PK Then
        PrefixToBusiness = BIZ_PK
    ElseIf left$(sku, Len(SKU_PREFIX_CL)) = SKU_PREFIX_CL Then
        PrefixToBusiness = BIZ_CL
    ElseIf left$(sku, Len(SKU_PREFIX_VG)) = SKU_PREFIX_VG Then
        PrefixToBusiness = BIZ_VG
    Else
        PrefixToBusiness = BIZ_VG
    End If
End Function

Private Function ActiveBusinessName() As String
    Dim ws As Worksheet, v As String
    v = CStr(GetConfigValue(CFG_ACTIVE_BIZ, BIZ_ALL))

    On Error Resume Next
    Set ws = Sheets("Dashboard")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If Trim$(CStr(ws.Range("B3").value)) <> "" Then v = CStr(ws.Range("B3").value)
    End If

    ActiveBusinessName = NormalizeBusiness(v)
End Function

' =====================================================================================
' HEADERS
' =====================================================================================
Public Function InventoryHeaders() As Variant
    InventoryHeaders = Array("SKU", "Business", "Console / Product", "Game Name", "Date Added", "Country", "Quality", "Quantity", "Paid", "From", "Seller", "Listed Status", "Listed Price", "Tested", "Bundle Number", "Stored", "Notes")
End Function

Public Function BatchEntryHeaders() As Variant
    BatchEntryHeaders = Array("Business", "SKU", "Console / Product", "Game Name", "Date Added", "Country", "Quality", "Quantity", "Paid", "From", "Seller", "Listed Status", "Listed Price", "Tested", "Bundle Number", "Stored", "Notes")
End Function

Private Sub GetInvCols(wsInv As Worksheet, ByRef c As InventoryCols)
    c.sku = MatchCol(wsInv, "SKU")
    c.Business = MatchCol(wsInv, "Business")
    c.Console = MatchConsoleColAny(wsInv)
    c.GameName = MatchCol(wsInv, "Game Name")
    c.DateAdded = MatchCol(wsInv, "Date Added")
    c.Country = MatchCol(wsInv, "Country")
    c.Quality = MatchCol(wsInv, "Quality")
    c.Quantity = MatchCol(wsInv, "Quantity")
    c.paid = MatchCol(wsInv, "Paid")
    c.FromCol = MatchCol(wsInv, "From")
    c.Seller = MatchCol(wsInv, "Seller")
    c.ListedStatus = MatchCol(wsInv, "Listed Status")
    c.ListedPrice = MatchCol(wsInv, "Listed Price")
    c.Tested = MatchCol(wsInv, "Tested")
    c.Bundle = MatchCol(wsInv, "Bundle Number")
    c.Stored = MatchCol(wsInv, "Stored")
    c.Notes = MatchCol(wsInv, "Notes")
End Sub

Private Sub EnsureInventoryHasHeadersAndBusiness()
    Dim wsInv As Worksheet, headers As Variant, cols As InventoryCols
    Dim lastCol As Long, lastRow As Long, r As Long

    Set wsInv = EnsureSheet("Inventory")
    ApplySheetChrome wsInv, False

    headers = InventoryHeaders()
    If Trim$(CStr(wsInv.Cells(1, 1).value)) = "" Then
        wsInv.Range(wsInv.Cells(1, 1), wsInv.Cells(1, UBound(headers) + 1)).value = headers
    End If

    GetInvCols wsInv, cols
    If cols.Business = 0 Then
        lastCol = wsInv.Cells(1, wsInv.Columns.Count).End(xlToLeft).Column
        wsInv.Cells(1, lastCol + 1).value = "Business"
        GetInvCols wsInv, cols
    End If

    lastRow = TableLastRow(wsInv, "A")
    For r = 2 To lastRow
        If cols.Business > 0 Then
            If Trim$(CStr(wsInv.Cells(r, cols.Business).value)) = "" Then
                wsInv.Cells(r, cols.Business).value = PrefixToBusiness(CStr(wsInv.Cells(r, cols.sku).value))
            End If
        End If
    Next r

    wsInv.Rows(1).Font.Bold = True
    wsInv.Rows(1).Interior.Color = CLR_PRIMARY
    wsInv.Rows(1).Font.Color = CLR_SECONDARY
End Sub

' =====================================================================================
' SKU GENERATION
' =====================================================================================
Public Function NextSKUForBusiness(ByVal biz As String) As String
    Dim key As String, prefix As String, n As Long
    biz = NormalizeBusiness(biz)
    prefix = BusinessToPrefix(biz)

    Select Case biz
        Case BIZ_PK: key = CFG_SKU_PK
        Case BIZ_CL: key = CFG_SKU_CL
        Case Else:   key = CFG_SKU_VG
    End Select

    n = CLng(Val(GetConfigValue(key, 0)))
    If n = 0 Then n = SeedLastSKUIndexForBusiness(biz)
    n = n + 1
    SaveConfigValue key, n
    NextSKUForBusiness = prefix & Format$(n, "000000")
End Function

Private Function SeedLastSKUIndexForBusiness(ByVal biz As String) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long, sku As String, n As Long, maxNum As Long
    Dim prefix As String, key As String

    biz = NormalizeBusiness(biz)
    prefix = BusinessToPrefix(biz)
    Select Case biz
        Case BIZ_PK: key = CFG_SKU_PK
        Case BIZ_CL: key = CFG_SKU_CL
        Case Else:   key = CFG_SKU_VG
    End Select

    Set ws = EnsureSheet("Inventory")
    lastRow = TableLastRow(ws, "A")
    For r = 2 To lastRow
        sku = CStr(ws.Cells(r, 1).value)
        If left$(sku, Len(prefix)) = prefix Then
            n = Val(Mid$(sku, Len(prefix) + 1))
            If n > maxNum Then maxNum = n
        End If
    Next r

    SaveConfigValue key, maxNum
    SeedLastSKUIndexForBusiness = maxNum
End Function

' =====================================================================================
' BUTTONS (SHAPES)
' =====================================================================================
Private Function AddAppButton( _
        ByVal ws As Worksheet, _
        ByVal btnCaption As String, _
        ByVal macroName As String, _
        ByVal left As Double, _
        ByVal top As Double, _
        Optional ByVal width As Double = 135, _
        Optional ByVal height As Double = 30, _
        Optional ByVal forcedName As String = "") As Shape

    EnsureThemeReady

    Dim shp As Shape, finalName As String
    finalName = IIf(forcedName = "", "btn_" & Replace(btnCaption, " ", "_") & "_" & CLng((Rnd + Timer) * 1000), forcedName)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, height)

    With shp
        On Error Resume Next: .Name = finalName: On Error GoTo 0
        .AlternativeText = macroName
        .OnAction = "HandleButtonPress"
        .Fill.ForeColor.RGB = CLR_BUTTON_BG
        .Line.ForeColor.RGB = CLR_BUTTON_BORDER
        .Shadow.Visible = msoFalse
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = btnCaption
            With .TextRange.Font
                .Name = THEME_FONT
                .Size = 10
                .Fill.ForeColor.RGB = CLR_BUTTON_TEXT
            End With
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Set AddAppButton = shp
End Function

Public Sub HandleButtonPress()
    Dim shp As Shape
    Dim macroToken As String
    Dim originalColor As Long
    Dim callerName As String

    On Error GoTo ErrHandler

    callerName = CStr(Application.Caller)
    Set shp = FindShapeByCaller(callerName)

    If shp Is Nothing Then
        MsgBox "Button pressed but shape not found: " & callerName, vbExclamation, MSG_TITLE
        Exit Sub
    End If

    macroToken = CStr(shp.AlternativeText)
    If Len(macroToken) = 0 Then
        MsgBox "Button has no macro token assigned (AlternativeText empty).", vbExclamation, MSG_TITLE
        Exit Sub
    End If

    originalColor = 0
    On Error Resume Next
    originalColor = shp.Fill.ForeColor.RGB
    shp.Fill.ForeColor.RGB = CLR_BUTTON_BG_HOVER
    shp.Line.ForeColor.RGB = CLR_ACCENT_ALT
    DoEvents
    On Error GoTo ErrHandler

    If InStr(1, macroToken, "|", vbTextCompare) > 0 Then
        HandleLookupAction_FromToken macroToken
        GoTo Restore
    End If

    Application.Run macroToken

Restore:
    On Error Resume Next
    If originalColor <> 0 Then shp.Fill.ForeColor.RGB = originalColor
    shp.Line.ForeColor.RGB = CLR_BUTTON_BORDER
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error in HandleButtonPress: " & Err.Number & " - " & Err.Description, vbCritical, MSG_TITLE
    On Error Resume Next
    If Not shp Is Nothing Then shp.Line.ForeColor.RGB = CLR_BUTTON_BORDER
End Sub

Private Function FindShapeByCaller(ByVal callerName As String) As Shape
    Dim ws As Worksheet, s As Shape

    On Error Resume Next
    If Not ActiveSheet Is Nothing Then
        Set s = ActiveSheet.Shapes(callerName)
        If Not s Is Nothing Then Set FindShapeByCaller = s: Exit Function
    End If
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set s = ws.Shapes(callerName)
        On Error GoTo 0
        If Not s Is Nothing Then Set FindShapeByCaller = s: Exit Function
    Next ws

    For Each ws In ThisWorkbook.Worksheets
        For Each s In ws.Shapes
            If InStr(1, s.Name, callerName, vbTextCompare) > 0 Or InStr(1, callerName, s.Name, vbTextCompare) > 0 Then
                Set FindShapeByCaller = s
                Exit Function
            End If
        Next s
    Next ws

    Set FindShapeByCaller = Nothing
End Function

Private Sub RecolorButtons(ws As Worksheet)
    Dim shp As Shape
    EnsureThemeReady
    On Error Resume Next
    For Each shp In ws.Shapes
        If left$(shp.Name, 4) = "btn_" Then
            shp.Fill.Visible = msoTrue
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = CLR_BUTTON_BG
            shp.Fill.Transparency = 0
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = CLR_BUTTON_BORDER
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_BUTTON_TEXT
        End If
    Next
    On Error GoTo 0
End Sub

' =====================================================================================
' SETTINGS (PER-BUSINESS NAMED RANGES)
' =====================================================================================
' Settings sheet schema:
' A = Business list
' B = Console / Product - Video Games
' C = Console / Product - Pokemon
' D = Console / Product - Cologne
' E = Quality - Video Games
' F = Quality - Pokemon
' G = Quality - Cologne
'
' Named ranges used by validation:
'   BusinessList
'   Console_Video_GamesList / Console_PokemonList / Console_CologneList
'   Quality_Video_GamesList / Quality_PokemonList / Quality_CologneList
'
' Plus, for your migration tool:
'   Field_console___product_Video_GamesList, etc (kept as-is)


Private Sub EnsureConsoleNamedRanges()
    ' Uses the SAME named ranges created by:
    ' RecreateBusinessAndPerBizNamedRanges / BuildCleanSettingsSchema
    '   Field_console___product_Video_GamesList
    '   Field_console___product_PokemonList
    '   Field_console___product_CologneList

    Dim ws As Worksheet, rb As Long, rC As Long, rD As Long
    Set ws = EnsureSheet("Settings")

    rb = Application.Max(2, ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
    rC = Application.Max(2, ws.Cells(ws.Rows.Count, "C").End(xlUp).Row)
    rD = Application.Max(2, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)

    On Error Resume Next
    ThisWorkbook.names("Field_console___product_Video_GamesList").Delete
    ThisWorkbook.names("Field_console___product_PokemonList").Delete
    ThisWorkbook.names("Field_console___product_CologneList").Delete
    On Error GoTo 0

    ThisWorkbook.names.Add Name:="Field_console___product_Video_GamesList", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(2, "B"), ws.Cells(rb, "B")).Address
    ThisWorkbook.names.Add Name:="Field_console___product_PokemonList", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(2, "C"), ws.Cells(rC, "C")).Address
    ThisWorkbook.names.Add Name:="Field_console___product_CologneList", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(2, "D"), ws.Cells(rD, "D")).Address
End Sub

Private Sub EnsureQualityNamedRanges()
    Dim ws As Worksheet, rE As Long, rF As Long, rG As Long
    Set ws = EnsureSheet("Settings")
    rE = Application.Max(2, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    rF = Application.Max(2, ws.Cells(ws.Rows.Count, "F").End(xlUp).Row)
    rG = Application.Max(2, ws.Cells(ws.Rows.Count, "G").End(xlUp).Row)

    On Error Resume Next
    ThisWorkbook.names("Quality_Video_GamesList").Delete
    ThisWorkbook.names("Quality_PokemonList").Delete
    ThisWorkbook.names("Quality_CologneList").Delete
    On Error GoTo 0

    ThisWorkbook.names.Add Name:="Quality_Video_GamesList", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(2, "E"), ws.Cells(rE, "E")).Address
    ThisWorkbook.names.Add Name:="Quality_PokemonList", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(2, "F"), ws.Cells(rF, "F")).Address
    ThisWorkbook.names.Add Name:="Quality_CologneList", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(2, "G"), ws.Cells(rG, "G")).Address
End Sub



' =====================================================================================
' INVENTORY VALIDATIONS (Console / Product + Quality linked to Business)
' =====================================================================================
Public Sub RebindInventoryDropdowns()
    SafeRun "Rebind Inventory Dropdowns", "RebindInventoryDropdowns_Worker"
End Sub

Private Sub RebindInventoryDropdowns_Worker()
    EnsureThemeReady
    EnsureSettingsBasics

    Dim wsInv As Worksheet, cols As InventoryCols
    Dim lastRow As Long, r As Long, bizVal As String
    Dim formulaCon As String, formulaQual As String

    Set wsInv = EnsureSheet("Inventory")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols
    If cols.Business = 0 Then Exit Sub

    lastRow = TableLastRow(wsInv, "A")
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        bizVal = NormalizeBusiness(CStr(wsInv.Cells(r, cols.Business).value))
        If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = PrefixToBusiness(CStr(wsInv.Cells(r, cols.sku).value))
        If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

        If cols.Console > 0 Then
            formulaCon = BusinessConsoleNamedList(bizVal)
            With wsInv.Cells(r, cols.Console).Validation
                .Delete
                .Add xlValidateList, xlValidAlertStop, xlBetween, formulaCon
            End With
        End If

        If cols.Quality > 0 Then
            formulaQual = BusinessQualityNamedList(bizVal)
            With wsInv.Cells(r, cols.Quality).Validation
                .Delete
                .Add xlValidateList, xlValidAlertStop, xlBetween, formulaQual
            End With
        End If
    Next r

    MsgBox "Inventory dropdowns rebound (Console/Product + Quality) based on each row's Business.", vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' ADD STOCK FORM (UI + DROPDOWN BINDING)
' =====================================================================================
Private Function AddStockField(ws As Worksheet, header As String) As Range
    Dim f As Range
    On Error Resume Next
    Set f = ws.Columns(1).Find(What:=header, LookAt:=xlWhole)
    On Error GoTo 0
    If Not f Is Nothing Then Set AddStockField = f.Offset(0, 1)
End Function

Private Function FieldVal(ws As Worksheet, header As String) As String
    Dim r As Range
    Set r = AddStockField(ws, header)
    If r Is Nothing Then FieldVal = "" Else FieldVal = Trim$(CStr(r.value))
End Function

Private Sub EnsureAddStockBusinessField(ws As Worksheet)
    Dim f As Range
    Set f = ws.Columns(1).Find(What:="Business", LookAt:=xlWhole)
    If f Is Nothing Then Exit Sub

    With f.Offset(0, 1).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=BusinessList"
    End With
    If Trim$(CStr(f.Offset(0, 1).value)) = "" Then
        Dim ab As String: ab = ActiveBusinessName()
        f.Offset(0, 1).value = IIf(ab = BIZ_ALL, BIZ_VG, ab)
    End If
End Sub

Public Sub SetupAddStockForm()
    Dim ws As Worksheet, labels As Variant, i As Long
    Set ws = EnsureSheet("Add Stock")
    ws.Cells.Clear
    ApplySheetChrome ws, False

    EnsureSettingsBasics

    labels = Array( _
        "SKU", "Business", "Console / Product", "Game Name", "Quality", "Quantity", "Date Added", "Country", _
        "Paid", "From", "Seller", "Listed Status", "Listed Price", "Tested", "Bundle Number", "Stored", "Notes", _
        "Barcodes (comma separated)" _
    )

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 35

    For i = LBound(labels) To UBound(labels)
        ws.Cells(i + 1, 1).value = labels(i)
        ws.Cells(i + 1, 1).Font.Bold = True
        ws.Cells(i + 1, 2).Interior.Color = CLR_INPUT
        ws.Cells(i + 1, 2).Borders.Color = CLR_BORDER
    Next i

    ' Force text for SKU + Barcode entry
    ws.Range("B1").NumberFormat = "@"
    ws.Cells(UBound(labels) + 1, 2).NumberFormat = "@"

    EnsureAddStockBusinessField ws
    BindAddStockConsoleDropdown ws
    BindAddStockQualityDropdown ws

    LayoutAddStockButtons
    RecolorButtons ws
End Sub

Private Sub LayoutAddStockButtons()
    Dim ws As Worksheet, shp As Shape
    Dim captions As Variant, macros As Variant
    Dim btnW As Double, btnH As Double, colGap As Double, rowGap As Double
    Dim baseLeft As Double, baseTop As Double
    Dim i As Long, r As Long, c As Long

    Set ws = EnsureSheet("Add Stock")

    For Each shp In ws.Shapes
        If left$(shp.Name, 4) = "btn_" Then shp.Delete
    Next shp

    captions = Array("Add Stock", "Remove Stock", "Attach Barcode", "Clear Entry", "Toggle Theme", "Rebind Dropdowns")
    macros = Array("AddStock_Single", "RemoveStock_Single", "AddBarcodeToSKU", "ClearAddStockForm", "ToggleDarkMode", "RebindAddStockDropdowns")

    btnW = 145: btnH = 30
    colGap = 12: rowGap = 12
    baseLeft = ws.Columns(5).left
    baseTop = ws.Rows(1).top + 2

    For i = LBound(captions) To UBound(captions)
        r = Int(i / 2)
        c = i Mod 2
        AddAppButton ws, captions(i), macros(i), baseLeft + c * (btnW + colGap), baseTop + r * (btnH + rowGap), btnW, btnH
    Next i
End Sub

Public Sub ClearAddStockForm()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Add Stock")
    ws.Range("B1:B200").ClearContents
    EnsureAddStockBusinessField ws
    BindAddStockConsoleDropdown ws
    BindAddStockQualityDropdown ws
End Sub

' =====================================================================================
' BARCODES (REGISTRY)
' =====================================================================================
Private Sub SetupBarcodesSheet(Optional ByVal ForceReset As Boolean = False)
    Dim ws As Worksheet
    Set ws = EnsureSheet("Barcodes")
    If ForceReset Then ws.Cells.Clear

    If ws.Cells(1, 1).value <> "Barcode" Then
        ws.Range("A1:C1").value = Array("Barcode", "SKU", "Barcode Count")
    End If

    ws.Columns("A").NumberFormat = "@"
    ApplySheetChrome ws, False
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = CLR_HEADER
End Sub

Private Function NormalizeBarcodeList(s As String) As String
    Dim t As String
    t = Trim$(s)
    If t = "" Then NormalizeBarcodeList = "": Exit Function
    t = Replace(t, vbCrLf, ",")
    t = Replace(t, vbLf, ",")
    t = Replace(t, vbCr, ",")
    t = Replace(t, ";", ",")
    t = Replace(t, "|", ",")
    Do While InStr(t, ",,") > 0: t = Replace(t, ",,", ","): Loop
    If left$(t, 1) = "," Then t = Mid$(t, 2)
    If Right$(t, 1) = "," Then t = left$(t, Len(t) - 1)
    NormalizeBarcodeList = t
End Function

Private Sub AddOrIncrementBarcode(wsBar As Worksheet, ByVal bc As String, ByVal sku As String)
    If Trim$(bc) = "" Or Trim$(sku) = "" Then Exit Sub
    bc = Trim$(bc)
    sku = Trim$(sku)

    Dim f As Range, firstAddr As String, done As Boolean
    wsBar.Columns(1).NumberFormat = "@"

    Set f = wsBar.Columns(1).Find(What:=bc, LookAt:=xlWhole)
    If Not f Is Nothing Then
        firstAddr = f.Address
        Do
            If Norm(CStr(f.value)) = Norm(bc) And Norm(CStr(f.Offset(0, 1).value)) = Norm(sku) Then
                f.Offset(0, 2).value = Application.Max(1, Val(f.Offset(0, 2).value) + 1)
                done = True
                Exit Do
            End If
            Set f = wsBar.Columns(1).FindNext(f)
            If f Is Nothing Or f.Address = firstAddr Then Exit Do
        Loop
    End If

    If Not done Then
        Dim lr As Long
        lr = wsBar.Cells(wsBar.Rows.Count, "A").End(xlUp).Row + 1
        wsBar.Cells(lr, 1).NumberFormat = "@"
        wsBar.Cells(lr, 1).value = bc
        wsBar.Cells(lr, 2).value = sku
        wsBar.Cells(lr, 3).value = 1
    End If
End Sub

Private Sub StoreBarcodes(ByVal sku As String, ByVal barcodesCsv As String)
    Dim wsBar As Worksheet, arr() As String, i As Long, normalized As String
    SetupBarcodesSheet False
    Set wsBar = EnsureSheet("Barcodes")

    normalized = NormalizeBarcodeList(barcodesCsv)
    If Trim$(normalized) = "" Then Exit Sub

    arr = Split(normalized, ",")
    For i = LBound(arr) To UBound(arr)
        If Trim$(arr(i)) <> "" Then AddOrIncrementBarcode wsBar, Trim$(arr(i)), sku
    Next i
End Sub

Private Function BarcodeCountForSKU(ByVal sku As String) As Long
    Dim wsBar As Worksheet, r As Long, lastRow As Long
    BarcodeCountForSKU = 0
    On Error Resume Next
    Set wsBar = Sheets("Barcodes")
    On Error GoTo 0
    If wsBar Is Nothing Then Exit Function

    lastRow = wsBar.Cells(wsBar.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If Norm(CStr(wsBar.Cells(r, 2).value)) = Norm(sku) Then BarcodeCountForSKU = BarcodeCountForSKU + 1
    Next r
End Function

Private Function ChooseSKUForBarcode(ByVal barcode As String, Optional ByVal preferBiz As String = "") As String
    Dim wsBar As Worksheet, wsInv As Worksheet, cols As InventoryCols
    Dim f As Range, firstAddr As String
    Dim matches As Collection
    Dim skuCandidate As String, bizFilter As String

    If Trim$(barcode) = "" Then Exit Function
    SetupBarcodesSheet False
    Set wsBar = EnsureSheet("Barcodes")

    Set wsInv = EnsureSheet("Inventory")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    bizFilter = NormalizeBusiness(preferBiz)
    If bizFilter = "" Then bizFilter = ActiveBusinessName()

    Set f = wsBar.Columns(1).Find(What:=barcode, LookAt:=xlWhole)
    If f Is Nothing Then
        If IsNumeric(barcode) Then
            Set f = wsBar.Columns(1).Find(What:=CStr(Format$(Val(barcode), "0")), LookAt:=xlWhole)
        End If
    End If
    If f Is Nothing Then Exit Function

    Set matches = New Collection
    firstAddr = f.Address
    Do
        skuCandidate = CStr(f.Offset(0, 1).value)
        If skuCandidate <> "" Then
            Dim invR As Range, rowBiz As String
            Set invR = wsInv.Columns(1).Find(What:=skuCandidate, LookAt:=xlWhole)
            If Not invR Is Nothing Then
                rowBiz = GetRowBusiness(wsInv, invR.Row, cols, True)
                If bizFilter = BIZ_ALL Or rowBiz = bizFilter Then
                    matches.Add f
                End If
            End If
        End If

        Set f = wsBar.Columns(1).FindNext(f)
        If f Is Nothing Or f.Address = firstAddr Then Exit Do
    Loop

    If matches.Count = 0 Then Exit Function
    If matches.Count = 1 Then
        ChooseSKUForBarcode = CStr(matches(1).Offset(0, 1).value)
        Exit Function
    End If

    Dim list As String, i As Long, sel As String, idx As Long
    list = "Multiple SKUs share barcode " & barcode & ":" & vbCrLf & vbCrLf
    For i = 1 To matches.Count
        list = list & i & ") " & CStr(matches(i).Offset(0, 1).value) & " (Count=" & CStr(matches(i).Offset(0, 2).value) & ")" & vbCrLf
    Next i
    list = list & vbCrLf & "Enter number to choose SKU (Cancel to skip):"
    sel = InputBox(list, MSG_TITLE)
    If sel = "" Then Exit Function
    If Not IsNumeric(sel) Then Exit Function
    idx = CLng(sel)
    If idx < 1 Or idx > matches.Count Then Exit Function

    ChooseSKUForBarcode = CStr(matches(idx).Offset(0, 1).value)
End Function

Public Sub AddBarcodeToSKU()
    Dim wsAdd As Worksheet, sku As String, bc As String
    Set wsAdd = EnsureSheet("Add Stock")

    sku = Trim$(FieldVal(wsAdd, "SKU"))
    If sku = "" Then
        sku = Trim$(InputBox("Enter SKU to attach barcode to:", MSG_TITLE))
        If sku = "" Then Exit Sub
    End If

    bc = Trim$(InputBox("Scan/enter barcode to attach to " & sku & ":", MSG_TITLE))
    If bc = "" Then Exit Sub

    StoreBarcodes sku, bc
    MsgBox "Attached barcode to " & sku, vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' INVENTORY CORE + HELPERS
' =====================================================================================
Private Function GetRowBusiness(ByVal wsInv As Worksheet, ByVal rowIdx As Long, ByRef cols As InventoryCols, Optional ByVal backfill As Boolean = True) As String
    Dim v As String
    If cols.Business > 0 Then v = Trim$(CStr(wsInv.Cells(rowIdx, cols.Business).value))
    If v = "" Then
        v = PrefixToBusiness(CStr(wsInv.Cells(rowIdx, cols.sku).value))
        If backfill And cols.Business > 0 Then
            On Error Resume Next
            wsInv.Cells(rowIdx, cols.Business).value = v
            On Error GoTo 0
        End If
    End If
    GetRowBusiness = NormalizeBusiness(v)
End Function

Private Function FindOrCreateInventoryRow(consoleVal As String, nameVal As String, qualityVal As String, qty As Long, ByRef outSKU As String, ByVal biz As String) As Long
    Dim wsInv As Worksheet, cols As InventoryCols, lastRow As Long, r As Long, rowBiz As String
    Set wsInv = EnsureSheet("Inventory")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols
    lastRow = TableLastRow(wsInv, "A")

    For r = 2 To lastRow
        If cols.GameName > 0 And cols.Console > 0 And cols.Quality > 0 Then
            If Norm(CStr(wsInv.Cells(r, cols.GameName).value)) = Norm(nameVal) _
               And Norm(CStr(wsInv.Cells(r, cols.Console).value)) = Norm(consoleVal) _
               And Norm(CStr(wsInv.Cells(r, cols.Quality).value)) = Norm(qualityVal) Then

                rowBiz = GetRowBusiness(wsInv, r, cols, True)
                If BusinessMatchesFilter(rowBiz, NormalizeBusiness(biz)) Then
                    wsInv.Cells(r, cols.Quantity).value = Val(wsInv.Cells(r, cols.Quantity).value) + qty
                    outSKU = CStr(wsInv.Cells(r, cols.sku).value)
                    FindOrCreateInventoryRow = r
                    Exit Function
                End If
            End If
        End If
    Next r

    outSKU = NextSKUForBusiness(biz)
    FindOrCreateInventoryRow = lastRow + 1
    wsInv.Cells(FindOrCreateInventoryRow, cols.sku).value = outSKU
    wsInv.Cells(FindOrCreateInventoryRow, cols.Business).value = NormalizeBusiness(biz)
    wsInv.Cells(FindOrCreateInventoryRow, cols.Console).value = consoleVal
    wsInv.Cells(FindOrCreateInventoryRow, cols.GameName).value = nameVal
    wsInv.Cells(FindOrCreateInventoryRow, cols.Quality).value = qualityVal
    wsInv.Cells(FindOrCreateInventoryRow, cols.Quantity).value = qty
End Function

Public Sub FillAddStockFromSKUIfNeeded(wsAdd As Worksheet, ByVal skuForm As String)
    On Error GoTo ErrHandler
    If Trim$(skuForm) = "" Then Exit Sub

    Dim wsInv As Worksheet: Set wsInv = EnsureSheet("Inventory")
    EnsureInventoryHasHeadersAndBusiness

    Dim cols As InventoryCols: GetInvCols wsInv, cols
    Dim f As Range: Set f = wsInv.Columns(cols.sku).Find(What:=skuForm, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Exit Sub

    Dim r As Long: r = f.Row
    Dim t As Range, v As String

    v = Trim$(CStr(wsInv.Cells(r, cols.Business).value))
    Set t = AddStockField(wsAdd, "Business"): If Not t Is Nothing Then If Trim$(CStr(t.value)) = "" Then t.value = v

    v = Trim$(CStr(wsInv.Cells(r, cols.Console).value))
    Set t = AddStockField(wsAdd, "Console / Product"): If Not t Is Nothing Then If Trim$(CStr(t.value)) = "" Then t.value = v

    v = Trim$(CStr(wsInv.Cells(r, cols.GameName).value))
    Set t = AddStockField(wsAdd, "Game Name"): If Not t Is Nothing Then If Trim$(CStr(t.value)) = "" Then t.value = v

    v = Trim$(CStr(wsInv.Cells(r, cols.Quality).value))
    Set t = AddStockField(wsAdd, "Quality"): If Not t Is Nothing Then If Trim$(CStr(t.value)) = "" Then t.value = v

    Exit Sub
ErrHandler:
    Resume Next
End Sub

' =====================================================================================
' ADD STOCK (SINGLE)
' =====================================================================================
Public Sub AddStock_Single(): SafeRun "Add Stock (Single)", "AddStock_Single_Worker": End Sub

Private Sub AddStock_Single_Worker()
    Dim wsInv As Worksheet, wsAdd As Worksheet, wsHist As Worksheet
    Dim cols As InventoryCols
    Dim qty As Long, skuForm As String, sku As String
    Dim rowTarget As Long, existing As Boolean
    Dim barcodes As String, addDateVal As Variant
    Dim f As Range, msg As String
    Dim conIn As String, nameIn As String, qualIn As String, bizIn As String
    Dim mismatch As String
    Dim arr() As String, i As Long, bc As String
    Dim wsBar As Worksheet, pickSKU As String, invR As Range

    Set wsInv = EnsureSheet("Inventory")
    Set wsAdd = EnsureSheet("Add Stock")
    Set wsHist = EnsureSheet("Stock History")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    qty = CLng(Val(FieldVal(wsAdd, "Quantity")))
    If qty <= 0 Then MsgBox "Enter a Quantity > 0.", vbExclamation, MSG_TITLE: Exit Sub

    skuForm = Trim$(FieldVal(wsAdd, "SKU"))
    conIn = Trim$(FieldVal(wsAdd, "Console / Product"))
    nameIn = Trim$(FieldVal(wsAdd, "Game Name"))
    qualIn = Trim$(FieldVal(wsAdd, "Quality"))
    bizIn = NormalizeBusiness(Trim$(FieldVal(wsAdd, "Business")))
    If bizIn = "" Or bizIn = BIZ_ALL Then
        Dim ab As String: ab = ActiveBusinessName()
        bizIn = IIf(ab = BIZ_ALL, BIZ_VG, ab)
        Dim tb As Range: Set tb = AddStockField(wsAdd, "Business")
        If Not tb Is Nothing Then If Trim$(CStr(tb.value)) = "" Then tb.value = bizIn
    End If

    ' If barcodes entered and no SKU, attempt to choose SKU based on barcode within business
    barcodes = Trim$(FieldVal(wsAdd, "Barcodes (comma separated)"))
    If skuForm = "" And barcodes <> "" Then
        SetupBarcodesSheet False
        Set wsBar = EnsureSheet("Barcodes")
        arr = Split(NormalizeBarcodeList(barcodes), ",")
        For i = LBound(arr) To UBound(arr)
            bc = Trim$(arr(i))
            If bc <> "" Then
                pickSKU = ChooseSKUForBarcode(bc, bizIn)
                If pickSKU <> "" Then
                    Set invR = wsInv.Columns(1).Find(What:=pickSKU, LookAt:=xlWhole)
                    If Not invR Is Nothing Then
                        If Norm(CStr(wsInv.Cells(invR.Row, cols.Quality).value)) = Norm(qualIn) Then
                            If nameIn <> "" And Norm(nameIn) <> Norm(CStr(wsInv.Cells(invR.Row, cols.GameName).value)) Then
                                If MsgBox("Barcode " & bc & " is linked to SKU " & pickSKU & " with Game Name '" & wsInv.Cells(invR.Row, cols.GameName).value & "'" & vbCrLf & _
                                   "Your form Game Name: '" & nameIn & "'" & vbCrLf & _
                                   "Add to existing SKU " & pickSKU & " instead of creating a new SKU?", vbYesNo + vbQuestion, MSG_TITLE) = vbYes Then
                                    skuForm = pickSKU
                                    Exit For
                                End If
                            Else
                                skuForm = pickSKU
                                Exit For
                            End If
                        Else
                            If MsgBox("Barcode " & bc & " links to SKU " & pickSKU & " but Quality differs." & vbCrLf & _
                               "Inventory Quality: '" & wsInv.Cells(invR.Row, cols.Quality).value & "'" & vbCrLf & _
                               "Form Quality: '" & qualIn & "'" & vbCrLf & _
                               "Do you still want to add to existing SKU " & pickSKU & "?", vbYesNo + vbExclamation, MSG_TITLE) = vbYes Then
                                skuForm = pickSKU
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If

    ' If only SKU provided, backfill form fields
    If skuForm <> "" And conIn = "" And nameIn = "" And qualIn = "" Then
        FillAddStockFromSKUIfNeeded wsAdd, skuForm
        conIn = Trim$(FieldVal(wsAdd, "Console / Product"))
        nameIn = Trim$(FieldVal(wsAdd, "Game Name"))
        qualIn = Trim$(FieldVal(wsAdd, "Quality"))
        bizIn = NormalizeBusiness(FieldVal(wsAdd, "Business"))
        If bizIn = "" Or bizIn = BIZ_ALL Then bizIn = PrefixToBusiness(skuForm)
    End If

    If skuForm = "" And nameIn = "" Then
        MsgBox "Game Name is required when adding stock without specifying an existing SKU.", vbExclamation, MSG_TITLE
        Exit Sub
    End If

    If skuForm <> "" Then
        Set f = wsInv.Columns(cols.sku).Find(What:=skuForm, LookAt:=xlWhole)
        If f Is Nothing Then MsgBox "SKU not found: " & skuForm, vbExclamation, MSG_TITLE: Exit Sub
        rowTarget = f.Row
        sku = skuForm
        existing = True

        ' Mismatch protection
        If conIn <> "" And Norm(conIn) <> Norm(CStr(wsInv.Cells(rowTarget, cols.Console).value)) Then _
            mismatch = mismatch & vbCrLf & "Console/Product mismatch: form=" & conIn & " inv=" & wsInv.Cells(rowTarget, cols.Console).value
        If nameIn <> "" And Norm(nameIn) <> Norm(CStr(wsInv.Cells(rowTarget, cols.GameName).value)) Then _
            mismatch = mismatch & vbCrLf & "Game Name mismatch: form=" & nameIn & " inv=" & wsInv.Cells(rowTarget, cols.GameName).value
        If qualIn <> "" And Norm(qualIn) <> Norm(CStr(wsInv.Cells(rowTarget, cols.Quality).value)) Then _
            mismatch = mismatch & vbCrLf & "Quality mismatch: form=" & qualIn & " inv=" & wsInv.Cells(rowTarget, cols.Quality).value

        If mismatch <> "" Then
            MsgBox "SKU details do not match:" & mismatch & vbCrLf & "No quantity added.", vbCritical, MSG_TITLE
            Exit Sub
        End If

        wsInv.Cells(rowTarget, cols.Quantity).value = Val(wsInv.Cells(rowTarget, cols.Quantity).value) + qty
        If Trim$(CStr(wsInv.Cells(rowTarget, cols.Business).value)) = "" Then wsInv.Cells(rowTarget, cols.Business).value = PrefixToBusiness(sku)

    Else
        rowTarget = FindOrCreateInventoryRow(conIn, nameIn, qualIn, qty, sku, bizIn)
        existing = False
    End If

    ' Date Added default
    If cols.DateAdded > 0 Then
        If Trim$(CStr(wsInv.Cells(rowTarget, cols.DateAdded).value)) = "" Then
            addDateVal = FieldVal(wsAdd, "Date Added")
            If addDateVal = "" Then addDateVal = Date
            wsInv.Cells(rowTarget, cols.DateAdded).value = addDateVal
        End If
    End If

    ' Set other blanks (non-destructive)
    If cols.Country > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.Country).value)) = "" Then wsInv.Cells(rowTarget, cols.Country).value = FieldVal(wsAdd, "Country")
    If cols.paid > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.paid).value)) = "" Then wsInv.Cells(rowTarget, cols.paid).value = FieldVal(wsAdd, "Paid")
    If cols.FromCol > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.FromCol).value)) = "" Then wsInv.Cells(rowTarget, cols.FromCol).value = FieldVal(wsAdd, "From")
    If cols.Seller > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.Seller).value)) = "" Then wsInv.Cells(rowTarget, cols.Seller).value = FieldVal(wsAdd, "Seller")
    If cols.ListedStatus > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.ListedStatus).value)) = "" Then wsInv.Cells(rowTarget, cols.ListedStatus).value = FieldVal(wsAdd, "Listed Status")
    If cols.ListedPrice > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.ListedPrice).value)) = "" Then wsInv.Cells(rowTarget, cols.ListedPrice).value = FieldVal(wsAdd, "Listed Price")
    If cols.Tested > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.Tested).value)) = "" Then wsInv.Cells(rowTarget, cols.Tested).value = FieldVal(wsAdd, "Tested")
    If cols.Bundle > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.Bundle).value)) = "" Then wsInv.Cells(rowTarget, cols.Bundle).value = FieldVal(wsAdd, "Bundle Number")
    If cols.Stored > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.Stored).value)) = "" Then wsInv.Cells(rowTarget, cols.Stored).value = FieldVal(wsAdd, "Stored")
    If cols.Notes > 0 And Trim$(CStr(wsInv.Cells(rowTarget, cols.Notes).value)) = "" Then wsInv.Cells(rowTarget, cols.Notes).value = FieldVal(wsAdd, "Notes")

    ' Store barcodes
    barcodes = FieldVal(wsAdd, "Barcodes (comma separated)")
    If barcodes <> "" Then StoreBarcodes CStr(sku), barcodes

    LogHistory_AddRemove "Add", wsInv, wsHist, rowTarget, qty, ""

    If existing Then
        msg = qty & " added to existing SKU " & sku & ". New Qty: " & wsInv.Cells(rowTarget, cols.Quantity).value
    Else
        msg = qty & " added under new SKU " & sku
    End If

    RefreshDashboard_Worker
    RepairInventoryColors
    MsgBox msg, vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' REMOVE / SOLD
' =====================================================================================
Public Sub RemoveStock_Single(): SafeRun "Remove Stock (Single)", "RemoveStock_Single_Worker": End Sub

Private Sub RemoveStock_Single_Worker()
    Dim wsInv As Worksheet, wsAdd As Worksheet, wsHist As Worksheet
    Dim qty As Long, item As String, sku As String, cols As InventoryCols
    Dim foundCell As Range
    Dim currentQty As Long, newQty As Long
    Dim markSold As VbMsgBoxResult
    Dim salePriceInput As String
    Dim salePrice As Double
    Dim totalCost As Double, realizedProfit As Double
    Dim qtyToProcess As Long, perUnit As Double
    Dim bizIn As String, ab As String

    Set wsInv = EnsureSheet("Inventory")
    Set wsAdd = EnsureSheet("Add Stock")
    Set wsHist = EnsureSheet("Stock History")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    sku = FieldVal(wsAdd, "SKU")
    item = FieldVal(wsAdd, "Game Name")
    qty = CLng(Val(FieldVal(wsAdd, "Quantity")))
    bizIn = NormalizeBusiness(FieldVal(wsAdd, "Business"))
    ab = ActiveBusinessName()
    If bizIn = "" Or bizIn = BIZ_ALL Then bizIn = IIf(ab = BIZ_ALL, BIZ_VG, ab)

    If (sku = "" And item = "") Or qty <= 0 Then MsgBox "Enter a SKU or Game Name and quantity > 0.", vbExclamation, MSG_TITLE: Exit Sub

    If sku <> "" Then Set foundCell = wsInv.Columns(cols.sku).Find(What:=sku, LookAt:=xlWhole)
    If foundCell Is Nothing And item <> "" Then
        Dim lastInvRow As Long, r As Long
        lastInvRow = TableLastRow(wsInv, "A")
        For r = 2 To lastInvRow
            If Norm(CStr(wsInv.Cells(r, cols.GameName).value)) = Norm(item) Then
                If BusinessMatchesFilter(GetRowBusiness(wsInv, r, cols, True), bizIn) Then
                    Set foundCell = wsInv.Cells(r, cols.GameName)
                    Exit For
                End If
            End If
        Next r
    End If

    If foundCell Is Nothing Then MsgBox "Item not found in selected business.", vbExclamation, MSG_TITLE: Exit Sub

    currentQty = Val(wsInv.Cells(foundCell.Row, cols.Quantity).value)
    If currentQty <= 0 Then MsgBox "Current quantity is zero; nothing to remove.", vbInformation, MSG_TITLE: Exit Sub

    If qty > currentQty Then qty = currentQty
    qtyToProcess = qty
    markSold = MsgBox("Treat this removal of " & qtyToProcess & " unit(s) as SOLD?", vbYesNo + vbQuestion, MSG_TITLE)

    If markSold = vbYes Then
        Dim suggestedTotal As Double
        suggestedTotal = Val(wsInv.Cells(foundCell.Row, cols.ListedPrice).value) * qtyToProcess
        salePriceInput = InputBox("Enter TOTAL sale price for these " & qtyToProcess & " unit(s)." & vbCrLf & "(Leave blank to record without price)", MSG_TITLE, IIf(suggestedTotal > 0, CStr(suggestedTotal), ""))
        If Trim$(salePriceInput) <> "" Then salePrice = Val(salePriceInput)

        newQty = currentQty - qtyToProcess
        wsInv.Cells(foundCell.Row, cols.Quantity).value = newQty

        If newQty = 0 And cols.ListedStatus > 0 Then wsInv.Cells(foundCell.Row, cols.ListedStatus).value = "Sold"

        If salePrice > 0 Then
            totalCost = Val(wsInv.Cells(foundCell.Row, cols.paid).value) * qtyToProcess
            realizedProfit = salePrice - totalCost
            perUnit = salePrice / qtyToProcess
        End If

        If salePrice > 0 And cols.Notes > 0 Then
            wsInv.Cells(foundCell.Row, cols.Notes).value = AggregateSaleNotes(CStr(wsInv.Cells(foundCell.Row, cols.Notes).value), qtyToProcess, salePrice, realizedProfit)
        End If

        LogHistory_AddRemove "Sold", wsInv, wsHist, foundCell.Row, -qtyToProcess, "", salePrice
        If sku = "" Then sku = CStr(wsInv.Cells(foundCell.Row, cols.sku).value)
        DecrementBarcodeCountsForSKU sku, qtyToProcess

        MsgBox "Recorded sale of " & qtyToProcess & " unit(s)." & IIf(salePrice > 0, vbCrLf & "Per Unit: " & FormatCurrency(perUnit) & vbCrLf & "Total Sale: " & FormatCurrency(salePrice) & vbCrLf & "Profit: " & FormatCurrency(realizedProfit), ""), vbInformation, MSG_TITLE
    Else
        newQty = currentQty - qtyToProcess
        If newQty < 0 Then newQty = 0
        wsInv.Cells(foundCell.Row, cols.Quantity).value = newQty
        LogHistory_AddRemove "Remove", wsInv, wsHist, foundCell.Row, -qtyToProcess, ""
        If sku = "" Then sku = CStr(wsInv.Cells(foundCell.Row, cols.sku).value)
        DecrementBarcodeCountsForSKU sku, qtyToProcess
        MsgBox qtyToProcess & " removed. New stock: " & newQty, vbInformation, MSG_TITLE
    End If

    RefreshDashboard_Worker
    ClearAddStockForm
End Sub

Private Sub DecrementBarcodeCountForPair(barcode As String, sku As String, qtySub As Long)
    If qtySub <= 0 Or Trim$(barcode) = "" Or Trim$(sku) = "" Then Exit Sub
    SetupBarcodesSheet False
    Dim wsBar As Worksheet, f As Range, firstAddr As String
    Set wsBar = EnsureSheet("Barcodes")
    Set f = wsBar.Columns(1).Find(What:=barcode, LookAt:=xlWhole)
    If f Is Nothing Then Exit Sub
    firstAddr = f.Address
    Do
        If Norm(CStr(f.Offset(0, 1).value)) = Norm(sku) Then
            f.Offset(0, 2).value = Application.Max(0, Val(f.Offset(0, 2).value) - qtySub)
            Exit Do
        End If
        Set f = wsBar.Columns(1).FindNext(f)
        If f Is Nothing Or f.Address = firstAddr Then Exit Do
    Loop
End Sub

Private Sub DecrementBarcodeCountsForSKU(sku As String, qtySub As Long, Optional preferBarcode As String = "")
    If qtySub <= 0 Then Exit Sub
    SetupBarcodesSheet False
    Dim wsBar As Worksheet, r As Long, lastRow As Long
    Dim matches As Collection, prompt As String, i As Long, sel As String, idx As Long

    Set wsBar = EnsureSheet("Barcodes")
    lastRow = wsBar.Cells(wsBar.Rows.Count, "A").End(xlUp).Row
    Set matches = New Collection

    For r = 2 To lastRow
        If Norm(CStr(wsBar.Cells(r, 2).value)) = Norm(sku) Then matches.Add r
    Next

    If matches.Count = 0 Then Exit Sub

    If Trim$(preferBarcode) <> "" Then
        DecrementBarcodeCountForPair preferBarcode, sku, qtySub
        Exit Sub
    End If

    If matches.Count = 1 Then
        r = matches(1)
        wsBar.Cells(r, 3).value = Application.Max(0, Val(wsBar.Cells(r, 3).value) - qtySub)
    Else
        prompt = "Multiple barcodes for SKU " & sku & ". Decrement which? Qty " & qtySub & vbCrLf
        For i = 1 To matches.Count
            r = matches(i)
            prompt = prompt & i & ") " & CStr(wsBar.Cells(r, 1).value) & " (Count=" & CStr(wsBar.Cells(r, 3).value) & ")" & vbCrLf
        Next i
        sel = InputBox(prompt, MSG_TITLE)
        If sel = "" Then Exit Sub
        If Not IsNumeric(sel) Then Exit Sub
        idx = CLng(sel)
        If idx < 1 Or idx > matches.Count Then Exit Sub
        r = matches(idx)
        wsBar.Cells(r, 3).value = Application.Max(0, Val(wsBar.Cells(r, 3).value) - qtySub)
    End If
End Sub

' =====================================================================================
' STOCK HISTORY
' =====================================================================================
Private Sub EnsureHistoryHeader(wsHist As Worksheet)
    Dim headers As Variant, i As Long
    If wsHist.Cells(1, 2).value = "" Then
        wsHist.Range("B1:H1").value = Array("Timestamp", "Action", "SKU", "Scanned Barcode", "Quantity Change", "New Total", "Sale Price")
        headers = InventoryHeaders()
        For i = LBound(headers) To UBound(headers)
            wsHist.Cells(1, 9 + i).value = headers(i)
        Next
    End If
    StyleStockHistoryHeader wsHist
End Sub

Private Sub StyleStockHistoryHeader(wsHist As Worksheet)
    EnsureThemeReady
    With wsHist.Range("B1:H1")
        .Font.Bold = True
        .Interior.Color = CLR_HISTORY_HEADER_BG
        .Font.Color = CLR_HISTORY_HEADER_TEXT
    End With
End Sub

Private Sub LogHistory_AddRemove(ByVal actionType As String, wsInv As Worksheet, wsHist As Worksheet, invRow As Long, qtyDelta As Long, Optional scannedBC As String = "", Optional salePrice As Variant)
    Dim r As Long, qtyCol As Long, headers As Variant, i As Long
    EnsureHistoryHeader wsHist

    r = wsHist.Cells(wsHist.Rows.Count, "B").End(xlUp).Row + 1
    wsHist.Cells(r, 2).value = Now
    wsHist.Cells(r, 3).value = actionType
    wsHist.Cells(r, 4).value = wsInv.Cells(invRow, 1).value
    wsHist.Cells(r, 5).NumberFormat = "@"
    wsHist.Cells(r, 5).value = scannedBC
    wsHist.Cells(r, 6).value = qtyDelta
    qtyCol = MatchCol(wsInv, "Quantity")
    If qtyCol > 0 Then wsHist.Cells(r, 7).value = wsInv.Cells(invRow, qtyCol).value

    If LCase$(actionType) = "sold" Then
        If IsMissing(salePrice) Then wsHist.Cells(r, 8).value = "" Else wsHist.Cells(r, 8).value = salePrice
    End If

    headers = InventoryHeaders()
    For i = LBound(headers) To UBound(headers)
        wsHist.Cells(r, 9 + i).value = wsInv.Cells(invRow, MatchCol(wsInv, headers(i))).value
    Next i

    Select Case LCase$(actionType)
        Case "add": wsHist.Rows(r).Interior.Color = CLR_SUCCESS
        Case "remove": wsHist.Rows(r).Interior.Color = CLR_ERROR
        Case "sold": wsHist.Rows(r).Interior.Color = CLR_NEUTRAL
    End Select
End Sub

Public Sub SetupStockHistoryButtons()
    Dim ws As Worksheet, btnNames As Variant, btnMacros As Variant, i As Long, topPos As Double, shp As Shape
    Set ws = EnsureSheet("Stock History")
    ApplySheetChrome ws, False

    For Each shp In ws.Shapes
        If left$(shp.Name, 4) = "btn_" Then shp.Delete
    Next

    ws.Columns("A").ColumnWidth = 18
    btnNames = Array("Reset History", "Export CSV", "Filter Adds", "Filter Removes", "Filter Sold", "Clear Filters", "Toggle Theme")
    btnMacros = Array("ConfirmResetStockHistory", "ExportStockHistory", "FilterAdds", "FilterRemoves", "FilterSold", "ClearFilters", "ToggleDarkMode")
    topPos = ws.Rows(2).top
    For i = LBound(btnNames) To UBound(btnNames)
        AddAppButton ws, btnNames(i), btnMacros(i), ws.Columns("A").left + 2, topPos + (i * 40), 120, 28
    Next
    EnsureHistoryHeader ws
End Sub

Public Sub ConfirmResetStockHistory()
    If MsgBox("Reset stock history (clear all)?", vbYesNo + vbQuestion, MSG_TITLE) = vbYes Then ResetStockHistory
End Sub

Public Sub ResetStockHistory()
    Dim wsHist As Worksheet
    Set wsHist = EnsureSheet("Stock History")
    wsHist.Cells.Clear
    EnsureHistoryHeader wsHist
    ApplySheetChrome wsHist, False
    SetupStockHistoryButtons
End Sub

Public Sub ExportStockHistory()
    Dim wsHist As Worksheet, filePath As String
    Set wsHist = EnsureSheet("Stock History")
    filePath = Application.GetSaveAsFilename("StockHistory.csv", "CSV Files (*.csv), *.csv")
    If filePath = "False" Then Exit Sub
    wsHist.Copy
    ActiveWorkbook.SaveAs Filename:=filePath, FileFormat:=xlCSV
    ActiveWorkbook.Close False
    MsgBox "Exported: " & filePath, vbInformation, MSG_TITLE
End Sub

Public Sub FilterAdds(): EnsureSheet("Stock History").Range("B1").CurrentRegion.AutoFilter Field:=2, Criteria1:="Add": End Sub
Public Sub FilterRemoves(): EnsureSheet("Stock History").Range("B1").CurrentRegion.AutoFilter Field:=2, Criteria1:="Remove": End Sub
Public Sub FilterSold(): EnsureSheet("Stock History").Range("B1").CurrentRegion.AutoFilter Field:=2, Criteria1:="Sold": End Sub
Public Sub ClearFilters()
    Dim ws As Worksheet: Set ws = EnsureSheet("Stock History")
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
End Sub

' =====================================================================================
' BARCODE LOOKUP
' =====================================================================================
Public Sub SetupBarcodeLookup()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Barcode Lookup")
    FormatLookupLayout ws
End Sub

Private Sub EnsureLookupBusinessSelector(wsLook As Worksheet)
    wsLook.Range("L2").value = "Active Business:"
    wsLook.Range("L2").Font.Bold = True
    With wsLook.Range("M2").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=BIZ_ALL & "," & BIZ_VG & "," & BIZ_PK & "," & BIZ_CL
    End With
    If Trim$(CStr(wsLook.Range("M2").value)) = "" Then wsLook.Range("M2").value = ActiveBusinessName()
End Sub

Private Sub RemoveLookupActionButtons(ws As Worksheet)
    Dim shp As Shape, toDelete As Collection, i As Long
    Set toDelete = New Collection
    For Each shp In ws.Shapes
        If LCase$(left$(shp.Name, 12)) = "btn_lookact_" Then
            toDelete.Add shp.Name
        ElseIf InStr(1, LCase$(CStr(shp.AlternativeText)), "lookupaction|", vbTextCompare) > 0 Then
            toDelete.Add shp.Name
        End If
    Next
    For i = 1 To toDelete.Count
        On Error Resume Next
        ws.Shapes(toDelete(i)).Delete
        On Error GoTo 0
    Next i
End Sub

Private Sub FormatLookupLayout(ws As Worksheet)
    ws.Cells.Clear
    ApplySheetChrome ws, False

    ws.Range("A2").value = "Search Mode:"
    ws.Range("A2").Font.Bold = True
    ws.Range("B2").value = "All"
    With ws.Range("B2").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Barcode,Name,SKU,All"
    End With

    ws.Range("A3").value = "Scan or Enter:"
    ws.Range("A3").Font.Bold = True
    With ws.Range("B3")
        .Interior.Color = CLR_INPUT
        .Borders.Color = CLR_BORDER
        .NumberFormat = "@"
    End With

    RemoveLookupActionButtons ws

    Dim leftStart As Double, topBtn As Double
    leftStart = ws.Columns("N").left
    topBtn = ws.Rows(2).top
    AddAppButton ws, "Lookup", "BarcodeLookup", leftStart, topBtn, 135, 30, "btnTop_Lookup"
    AddAppButton ws, "Toggle Theme", "ToggleDarkMode", leftStart + 150, topBtn, 135, 30, "btnTop_ToggleTheme"

    EnsureLookupBusinessSelector ws

    ws.Range("A5:J5").value = Array("SKU", "Business", "Console / Product", "Game Name", "Quality", "Country", "Quantity", "Listed Price", "Barcode Count", "MatchedBy")
    ws.Range("A5:J5").Font.Bold = True
    ws.Range("A5:J5").Interior.Color = CLR_HEADER

    RecolorButtons ws
End Sub

Public Sub BarcodeLookup(): SafeRun "Barcode Lookup", "BarcodeLookup_Worker": End Sub

Private Sub BarcodeLookup_Worker()
    EnsureThemeReady

    Dim wsInv As Worksheet, wsLook As Worksheet, wsBar As Worksheet
    Dim term As String, cols As InventoryCols, searchMode As String, rawB As Variant
    Dim dict As Object
    Dim bizFilter As String
    Dim lastBarRow As Long, f As Range, firstAddr As String
    Dim lastInvRow As Long, r As Long, nm As String, invR As Range, skuFound As String

    Set wsInv = EnsureSheet("Inventory")
    Set wsLook = EnsureSheet("Barcode Lookup")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    SetupBarcodesSheet False
    Set wsBar = EnsureSheet("Barcodes")

    searchMode = LCase$(Trim$(CStr(wsLook.Range("B2").value)))
    rawB = wsLook.Range("B3").value
    If IsError(rawB) Then Exit Sub
    term = Trim$(CStr(rawB))
    If term = "" Then Exit Sub

    bizFilter = NormalizeBusiness(CStr(wsLook.Range("M2").value))
    If bizFilter = "" Then bizFilter = ActiveBusinessName()
    SaveConfigValue CFG_ACTIVE_BIZ, bizFilter

    wsLook.Range("A6:Z2000").ClearContents
    RemoveLookupActionButtons wsLook

    wsLook.Range("A5:J5").value = Array("SKU", "Business", "Console / Product", "Game Name", "Quality", "Country", "Quantity", "Listed Price", "Barcode Count", "MatchedBy")
    wsLook.Range("A5:J5").Font.Bold = True
    wsLook.Range("A5:J5").Interior.Color = CLR_HEADER

    Set dict = CreateObject("Scripting.Dictionary")

    ' 1) Barcode search
    If searchMode = "barcode" Or searchMode = "all" Then
        lastBarRow = wsBar.Cells(wsBar.Rows.Count, "A").End(xlUp).Row
        If lastBarRow >= 2 Then
            Set f = wsBar.Columns(1).Find(What:=term, LookAt:=xlWhole)
            If f Is Nothing Then
                If IsNumeric(term) Then
                    Set f = wsBar.Columns(1).Find(What:=CStr(Format$(Val(term), "0")), LookAt:=xlWhole)
                End If
            End If

            If Not f Is Nothing Then
                firstAddr = f.Address
                Do
                    skuFound = CStr(f.Offset(0, 1).value)
                    If skuFound <> "" Then
                        Set invR = wsInv.Columns(1).Find(What:=skuFound, LookAt:=xlWhole)
                        If Not invR Is Nothing Then
                            Dim rb As String: rb = GetRowBusiness(wsInv, invR.Row, cols, True)
                            If BusinessMatchesFilter(rb, bizFilter) Then
                                If Not dict.Exists(skuFound) Then
                                    dict.Add skuFound, Array(rb, BarcodeCountForSKU(skuFound), "Barcode")
                                Else
                                    Dim a0 As Variant: a0 = dict(skuFound)
                                    If InStr(1, CStr(a0(2)), "Barcode", vbTextCompare) = 0 Then
                                        a0(2) = a0(2) & "+Barcode"
                                        dict(skuFound) = a0
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Set f = wsBar.Columns(1).FindNext(f)
                    If f Is Nothing Or f.Address = firstAddr Then Exit Do
                Loop
            End If
        End If
    End If

    ' 2) Name search
    If searchMode = "name" Or searchMode = "all" Then
        lastInvRow = TableLastRow(wsInv, "A")
        For r = 2 To lastInvRow
            nm = CStr(wsInv.Cells(r, cols.GameName).value)
            If Len(Trim$(nm)) > 0 Then
                If InStr(1, nm, term, vbTextCompare) > 0 Then
                    Dim rb2 As String: rb2 = GetRowBusiness(wsInv, r, cols, True)
                    If BusinessMatchesFilter(rb2, bizFilter) Then
                        skuFound = CStr(wsInv.Cells(r, cols.sku).value)
                        If skuFound <> "" Then
                            If Not dict.Exists(skuFound) Then
                                dict.Add skuFound, Array(rb2, BarcodeCountForSKU(skuFound), "Name")
                            Else
                                Dim a1 As Variant: a1 = dict(skuFound)
                                If InStr(1, CStr(a1(2)), "Name", vbTextCompare) = 0 Then
                                    a1(2) = a1(2) & "+Name"
                                    dict(skuFound) = a1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next r
    End If

    ' 3) SKU search
    If searchMode = "sku" Or searchMode = "all" Then
        Set invR = wsInv.Columns(1).Find(What:=term, LookAt:=xlWhole)
        If Not invR Is Nothing Then
            Dim rb3 As String: rb3 = GetRowBusiness(wsInv, invR.Row, cols, True)
            If BusinessMatchesFilter(rb3, bizFilter) Then
                skuFound = CStr(invR.value)
                If Not dict.Exists(skuFound) Then
                    dict.Add skuFound, Array(rb3, BarcodeCountForSKU(skuFound), "SKU")
                Else
                    Dim a2 As Variant: a2 = dict(skuFound)
                    If InStr(1, CStr(a2(2)), "SKU", vbTextCompare) = 0 Then
                        a2(2) = a2(2) & "+SKU"
                        dict(skuFound) = a2
                    End If
                End If
            End If
        ElseIf searchMode = "all" Then
            lastInvRow = TableLastRow(wsInv, "A")
            For r = 2 To lastInvRow
                If InStr(1, CStr(wsInv.Cells(r, 1).value), term, vbTextCompare) > 0 Then
                    Dim rb4 As String: rb4 = GetRowBusiness(wsInv, r, cols, True)
                    If BusinessMatchesFilter(rb4, bizFilter) Then
                        skuFound = CStr(wsInv.Cells(r, 1).value)
                        If Not dict.Exists(skuFound) Then
                            dict.Add skuFound, Array(rb4, BarcodeCountForSKU(skuFound), "SKU")
                        Else
                            Dim a3 As Variant: a3 = dict(skuFound)
                            If InStr(1, CStr(a3(2)), "SKU", vbTextCompare) = 0 Then
                                a3(2) = a3(2) & "+SKU"
                                dict(skuFound) = a3
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    End If

    If dict.Count = 0 Then
        wsLook.Range("A7").value = "No results for: " & term & " (Business: " & bizFilter & ")"
        wsLook.Range("A7").Font.Color = CLR_ERROR
        Exit Sub
    End If

    Dim outRow As Long: outRow = 7
    Dim k As Variant, dat
    For Each k In dict.Keys
        Set invR = wsInv.Columns(1).Find(What:=k, LookAt:=xlWhole)
        If Not invR Is Nothing Then
            dat = dict(k) ' (biz, barcodeCount, matchedBy)
            wsLook.Cells(outRow, 1).value = k
            wsLook.Cells(outRow, 2).value = dat(0)
            wsLook.Cells(outRow, 3).value = wsInv.Cells(invR.Row, cols.Console).value
            wsLook.Cells(outRow, 4).value = wsInv.Cells(invR.Row, cols.GameName).value
            wsLook.Cells(outRow, 5).value = wsInv.Cells(invR.Row, cols.Quality).value
            wsLook.Cells(outRow, 6).value = wsInv.Cells(invR.Row, cols.Country).value
            wsLook.Cells(outRow, 7).value = wsInv.Cells(invR.Row, cols.Quantity).value
            wsLook.Cells(outRow, 8).value = wsInv.Cells(invR.Row, cols.ListedPrice).value
            wsLook.Cells(outRow, 9).value = dat(1)
            wsLook.Cells(outRow, 10).value = dat(2)
            If outRow Mod 2 = 1 Then wsLook.Range("A" & outRow & ":J" & outRow).Interior.Color = CLR_ALTROW

            Dim leftX As Double, topY As Double
            Dim forcedNameSell As String, forcedNameRemove As String, forcedNameAdd As String

            leftX = wsLook.Columns("L").left + 2
            topY = wsLook.Rows(outRow).top + 2

            forcedNameSell = "btn_LookAct_Sell_" & Replace(CStr(k), " ", "_") & "_" & outRow
            forcedNameRemove = "btn_LookAct_Remove_" & Replace(CStr(k), " ", "_") & "_" & outRow
            forcedNameAdd = "btn_LookAct_Add_" & Replace(CStr(k), " ", "_") & "_" & outRow

            AddAppButton wsLook, "Sell", "LookupAction|Sell|" & k, leftX, topY, 60, 20, forcedNameSell
            AddAppButton wsLook, "Remove", "LookupAction|Remove|" & k, leftX + 70, topY, 60, 20, forcedNameRemove
            AddAppButton wsLook, "Add Qty", "LookupAction|Add|" & k, leftX + 140, topY, 60, 20, forcedNameAdd

            outRow = outRow + 1
        End If
    Next
    wsLook.Columns("A:J").AutoFit
End Sub

Private Sub HandleLookupAction_FromToken(ByVal token As String)
    Dim parts() As String
    parts = Split(token, "|")
    If UBound(parts) < 1 Then Exit Sub
    Dim action As String, sku As String
    action = "": sku = ""
    If UBound(parts) >= 1 Then action = parts(1)
    If UBound(parts) >= 2 Then sku = parts(2)
    If sku = "" Then
        sku = Trim$(InputBox("Enter SKU to " & LCase$(action) & ":", MSG_TITLE, ""))
        If sku = "" Then Exit Sub
    End If
    Select Case LCase$(action)
        Case "sell"
            PerformSellFromSKU sku
        Case "remove"
            PerformRemoveFromSKU sku
        Case "add"
            PerformAddFromSKU sku
        Case Else
            MsgBox "Unknown lookup action: " & action, vbExclamation, MSG_TITLE
    End Select
End Sub

Private Sub PerformAddFromSKU(ByVal sku As String)
    Dim wsInv As Worksheet, wsHist As Worksheet, cols As InventoryCols
    Dim rInv As Range, qtyToAdd As Long, currentQty As Long
    Dim addDate As Variant

    Set wsInv = EnsureSheet("Inventory")
    Set wsHist = EnsureSheet("Stock History")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    Set rInv = wsInv.Columns(1).Find(What:=sku, LookAt:=xlWhole)
    If rInv Is Nothing Then
        MsgBox "SKU not found: " & sku, vbExclamation, MSG_TITLE: Exit Sub
    End If

    currentQty = Val(wsInv.Cells(rInv.Row, cols.Quantity).value)
    qtyToAdd = CLng(Val(InputBox("Quantity to ADD for " & sku & ":", MSG_TITLE, 1)))
    If qtyToAdd <= 0 Then Exit Sub

    addDate = Trim$(InputBox("Date to record for these items (leave blank to skip, format mm/dd/yyyy):", MSG_TITLE, ""))
    On Error Resume Next
    If Not IsDate(addDate) Then addDate = Empty
    On Error GoTo 0

    wsInv.Cells(rInv.Row, cols.Quantity).value = currentQty + qtyToAdd
    If Not IsEmpty(addDate) Then
        If cols.DateAdded > 0 Then wsInv.Cells(rInv.Row, cols.DateAdded).value = CDate(addDate)
    End If

    LogHistory_AddRemove "Add", wsInv, wsHist, rInv.Row, qtyToAdd, ""
    RefreshDashboard_Worker
    RepairInventoryColors
End Sub

Private Sub PerformRemoveFromSKU(ByVal sku As String)
    Dim wsInv As Worksheet, wsHist As Worksheet, cols As InventoryCols
    Dim rInv As Range, qtyToRemove As Long, currentQty As Long
    Set wsInv = EnsureSheet("Inventory")
    Set wsHist = EnsureSheet("Stock History")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    Set rInv = wsInv.Columns(1).Find(What:=sku, LookAt:=xlWhole)
    If rInv Is Nothing Then MsgBox "SKU not found: " & sku, vbExclamation, MSG_TITLE: Exit Sub

    currentQty = Val(wsInv.Cells(rInv.Row, cols.Quantity).value)
    qtyToRemove = CLng(Val(InputBox("Quantity to remove:", MSG_TITLE, 1)))
    If qtyToRemove <= 0 Then Exit Sub
    If qtyToRemove > currentQty Then qtyToRemove = currentQty
    If MsgBox("Remove " & qtyToRemove & " units from " & sku & "?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub

    wsInv.Cells(rInv.Row, cols.Quantity).value = currentQty - qtyToRemove
    LogHistory_AddRemove "Remove", wsInv, wsHist, rInv.Row, -qtyToRemove, ""
    DecrementBarcodeCountsForSKU sku, qtyToRemove
    RefreshDashboard_Worker
    RepairInventoryColors
End Sub

Private Sub PerformSellFromSKU(ByVal sku As String)
    Dim wsInv As Worksheet, wsHist As Worksheet, cols As InventoryCols
    Dim rInv As Range, qtyToSell As Long, currentQty As Long
    Dim salePrice As Double, totalCost As Double, realizedProfit As Double, perUnit As Double
    Set wsInv = EnsureSheet("Inventory")
    Set wsHist = EnsureSheet("Stock History")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    Set rInv = wsInv.Columns(1).Find(What:=sku, LookAt:=xlWhole)
    If rInv Is Nothing Then MsgBox "SKU not found: " & sku, vbExclamation, MSG_TITLE: Exit Sub

    currentQty = Val(wsInv.Cells(rInv.Row, cols.Quantity).value)
    qtyToSell = CLng(Val(InputBox("Quantity sold:", MSG_TITLE, 1)))
    If qtyToSell <= 0 Then Exit Sub
    If qtyToSell > currentQty Then qtyToSell = currentQty

    salePrice = Val(InputBox("Sale Price (TOTAL for these " & qtyToSell & "):", MSG_TITLE, Val(wsInv.Cells(rInv.Row, cols.ListedPrice).value) * qtyToSell))

    wsInv.Cells(rInv.Row, cols.Quantity).value = currentQty - qtyToSell
    If currentQty - qtyToSell = 0 And cols.ListedStatus > 0 Then wsInv.Cells(rInv.Row, cols.ListedStatus).value = "Sold"

    totalCost = Val(wsInv.Cells(rInv.Row, cols.paid).value) * qtyToSell
    realizedProfit = salePrice - totalCost
    If qtyToSell > 0 Then perUnit = salePrice / qtyToSell

    If cols.Notes > 0 And salePrice > 0 Then
        wsInv.Cells(rInv.Row, cols.Notes).value = AggregateSaleNotes(CStr(wsInv.Cells(rInv.Row, cols.Notes).value), qtyToSell, salePrice, realizedProfit)
    End If

    LogHistory_AddRemove "Sold", wsInv, wsHist, rInv.Row, -qtyToSell, "", salePrice
    DecrementBarcodeCountsForSKU sku, qtyToSell
    RefreshDashboard_Worker
    RepairInventoryColors

    MsgBox "Sold " & qtyToSell & " of " & sku & vbCrLf & _
           "Per Unit: " & FormatCurrency(perUnit) & vbCrLf & _
           "Total Sale: " & FormatCurrency(salePrice) & vbCrLf & _
           "Profit: " & FormatCurrency(realizedProfit), vbInformation, MSG_TITLE
End Sub

Public Sub SetupBatchEntry()
    Dim ws As Worksheet, headers As Variant, i As Long, lastCol As Long
    Set ws = EnsureSheet("Batch Entry")
    ws.Cells.Clear
    ApplySheetChrome ws, False

    EnsureSettingsBasics

    headers = BatchEntryHeaders()
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Cells(1, i + 1).Interior.Color = CLR_HEADER
    Next i

    ws.Cells(1, UBound(headers) + 2).value = "Barcodes (comma separated)"
    ws.Cells(1, UBound(headers) + 2).Font.Bold = True
    ws.Cells(1, UBound(headers) + 2).Interior.Color = CLR_HEADER
    lastCol = UBound(headers) + 2

    ws.Range(ws.Cells(2, 1), ws.Cells(300, lastCol)).Borders.Color = CLR_BORDER
    ws.Range(ws.Cells(2, lastCol), ws.Cells(300, lastCol)).Interior.Color = CLR_INPUT
    ws.Columns(lastCol).ColumnWidth = 28

    LayoutBatchButtons
    ApplyBatchBusinessDropdown ws
    BindBatchConsoleDropdowns ws
    BindBatchQualityDropdowns ws
End Sub

Private Sub LayoutBatchButtons()
    Dim ws As Worksheet, btnNames As Variant, btnMacros As Variant, i As Long
    Dim topStart As Double, colRight As Long, leftPos As Double, shp As Shape

    Set ws = EnsureSheet("Batch Entry")
    For Each shp In ws.Shapes
        If left$(shp.Name, 4) = "btn_" Then shp.Delete
    Next

    btnNames = Array("Commit Batch", "Clear Batch", "Import File", "Toggle Theme", "Rebind Dropdowns")
    btnMacros = Array("CommitBatchEntry", "ClearBatchEntry", "ImportBatchFile", "ToggleDarkMode", "RebindBatchDropdowns")

    colRight = ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column).Column + 1
    leftPos = ws.Columns(colRight).left + 2
    ws.Columns(colRight).ColumnWidth = 16
    topStart = ws.Rows(1).top + 5

    For i = LBound(btnNames) To UBound(btnNames)
        AddAppButton ws, btnNames(i), btnMacros(i), leftPos, topStart + (i * 40), 140, 30
    Next
End Sub

Private Sub ApplyBatchBusinessDropdown(ws As Worksheet)
    Dim bizCol As Long
    bizCol = MatchCol(ws, "Business")
    If bizCol = 0 Then Exit Sub

    With ws.Range(ws.Cells(2, bizCol), ws.Cells(300, bizCol)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=BusinessList"
    End With
End Sub

Public Sub CommitBatchEntry(): SafeRun "Commit Batch Entry", "CommitBatchEntry_Worker": End Sub

Private Sub CommitBatchEntry_Worker()
    Dim wsBatch As Worksheet, wsInv As Worksheet, wsHist As Worksheet
    Dim cols As InventoryCols
    Dim lastRow As Long, r As Long
    Dim qty As Long, bizVal As String, sku As String
    Dim invRow As Long
    Dim colSku As Long, colBiz As Long, colQty As Long, colBar As Long

    Set wsBatch = EnsureSheet("Batch Entry")
    Set wsInv = EnsureSheet("Inventory")
    Set wsHist = EnsureSheet("Stock History")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    colSku = MatchCol(wsBatch, "SKU")
    colBiz = MatchCol(wsBatch, "Business")
    colQty = MatchCol(wsBatch, "Quantity")
    colBar = MatchCol(wsBatch, "Barcodes (comma separated)")
    If colBiz = 0 Or colQty = 0 Then
        MsgBox "Batch Entry headers missing (need Business and Quantity).", vbCritical, MSG_TITLE
        Exit Sub
    End If

    lastRow = wsBatch.Cells(wsBatch.Rows.Count, colBiz).End(xlUp).Row
    If lastRow < 2 Then MsgBox "No batch data.", vbExclamation, MSG_TITLE: Exit Sub

    For r = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsBatch.Rows(r)) = 0 Then GoTo NextRow

        sku = IIf(colSku > 0, Trim$(CStr(wsBatch.Cells(r, colSku).value)), "")
        qty = CLng(Val(wsBatch.Cells(r, colQty).value))
        If qty <= 0 Then qty = 1

        bizVal = NormalizeBusiness(CStr(wsBatch.Cells(r, colBiz).value))
        If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

        invRow = CommitOneBatchRow(wsBatch, r, wsInv, cols, sku, qty, bizVal)

        If colBar > 0 Then
            Dim barcodes As String
            barcodes = CStr(wsBatch.Cells(r, colBar).value)
            If Trim$(barcodes) <> "" Then StoreBarcodes CStr(wsInv.Cells(invRow, cols.sku).value), barcodes
        End If

        LogHistory_AddRemove "Add", wsInv, wsHist, invRow, qty, ""
NextRow:
    Next r

    RefreshDashboard_Worker
    RepairInventoryColors
    MsgBox "Batch committed.", vbInformation, MSG_TITLE
End Sub

Private Function CommitOneBatchRow(wsBatch As Worksheet, ByVal batchRow As Long, wsInv As Worksheet, ByRef cols As InventoryCols, ByVal sku As String, ByVal qty As Long, ByVal bizVal As String) As Long
    Dim f As Range, invRow As Long

    ' If SKU exists: increment qty
    If sku <> "" Then
        Set f = wsInv.Columns(cols.sku).Find(What:=sku, LookAt:=xlWhole)
        If Not f Is Nothing Then
            invRow = f.Row
            wsInv.Cells(invRow, cols.Quantity).value = Val(wsInv.Cells(invRow, cols.Quantity).value) + qty
            If Trim$(CStr(wsInv.Cells(invRow, cols.Business).value)) = "" Then wsInv.Cells(invRow, cols.Business).value = PrefixToBusiness(sku)
            CommitOneBatchRow = invRow
            Exit Function
        End If
    End If

    ' Create row
    invRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row + 1
    If invRow < 2 Then invRow = 2
    If sku = "" Then sku = NextSKUForBusiness(bizVal)

    wsInv.Cells(invRow, cols.sku).value = sku
    wsInv.Cells(invRow, cols.Business).value = bizVal
    wsInv.Cells(invRow, cols.Quantity).value = qty

    ' Map by header name where possible (excluding Quantity/SKU/Business)
    Dim h As Variant, invCol As Long, batchCol As Long
    For Each h In InventoryHeaders()
        invCol = MatchCol(wsInv, h)
        batchCol = MatchCol(wsBatch, h)
        If invCol > 0 And batchCol > 0 Then
            If h <> "SKU" And h <> "Quantity" And h <> "Business" Then
                If Trim$(CStr(wsInv.Cells(invRow, invCol).value)) = "" Then
                    wsInv.Cells(invRow, invCol).value = wsBatch.Cells(batchRow, batchCol).value
                End If
            End If
        End If
    Next h

    CommitOneBatchRow = invRow
End Function

Public Sub ClearBatchEntry()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Batch Entry")
    ws.Range("A2:Z2000").ClearContents
    RebindBatchDropdowns
    MsgBox "Batch cleared.", vbInformation, MSG_TITLE
End Sub

Public Sub ImportBatchFile()
    MsgBox "ImportBatchFile not implemented in this build.", vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' NOTES AGGREGATION FOR SOLD
' =====================================================================================
Private Function ParseNumberAfter(ByVal txt As String, ByVal key As String) As Long
    Dim p As Long, s As String, i As Long, ch As String
    p = InStr(1, LCase$(txt), LCase$(key))
    If p = 0 Then Exit Function
    p = p + Len(key)
    Do While p <= Len(txt) And Mid$(txt, p, 1) = " "
        p = p + 1
    Loop
    For i = p To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then
            s = s & ch
        ElseIf s <> "" Then
            Exit For
        End If
    Next i
    If s <> "" Then ParseNumberAfter = CLng(Val(s))
End Function

Private Function ParseCurrencyBetween(ByVal txt As String, ByVal marker As String) As Double
    Dim p As Long, i As Long, ch As String, started As Boolean, s As String
    p = InStr(1, LCase$(txt), LCase$(marker))
    If p = 0 Then Exit Function
    p = p + Len(marker)
    Do While p <= Len(txt) And (Mid$(txt, p, 1) = " " Or Mid$(txt, p, 1) = "$")
        p = p + 1
    Loop
    For i = p To Len(txt)
        ch = Mid$(txt, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "," Then
            s = s & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i
    If s <> "" Then ParseCurrencyBetween = Val(Replace(s, ",", ""))
End Function

Private Function AggregateSaleNotes(existingNotes As String, newQty As Long, newSaleTotal As Double, newProfit As Double) As String
    Dim parts() As String
    Dim i As Long, token As String
    Dim otherParts As Collection, saleDict As Object
    Dim priceKey As String, qty As Long, unitPrice As Double, totalProfit As Double
    Dim perUnit As Double, rebuilt As String

    Set otherParts = New Collection
    Set saleDict = CreateObject("Scripting.Dictionary")

    If Trim$(existingNotes) <> "" Then
        parts = Split(existingNotes, " | ")
        For i = LBound(parts) To UBound(parts)
            token = Trim$(parts(i))
            If LCase$(left$(token, 4)) = "sold" Then
                qty = ParseNumberAfter(token, "Sold")
                unitPrice = ParseCurrencyBetween(token, "@Total=")
                totalProfit = ParseCurrencyBetween(token, "Profit=")
                If qty > 0 Then
                    priceKey = Format$(unitPrice, "0.00")
                    If Not saleDict.Exists(priceKey) Then
                        saleDict.Add priceKey, Array(qty, unitPrice, totalProfit)
                    Else
                        Dim arrExisting As Variant
                        arrExisting = saleDict(priceKey)
                        arrExisting(0) = arrExisting(0) + qty
                        arrExisting(2) = arrExisting(2) + totalProfit
                        saleDict(priceKey) = arrExisting
                    End If
                Else
                    otherParts.Add token
                End If
            Else
                otherParts.Add token
            End If
        Next i
    End If

    If newQty > 0 Then
        perUnit = newSaleTotal / newQty
        priceKey = Format$(perUnit, "0.00")
        If saleDict.Exists(priceKey) Then
            Dim arr2 As Variant
            arr2 = saleDict(priceKey)
            arr2(0) = arr2(0) + newQty
            arr2(2) = arr2(2) + newProfit
            saleDict(priceKey) = arr2
        Else
            saleDict.Add priceKey, Array(newQty, perUnit, newProfit)
        End If
    End If

    Dim outParts As Collection
    Set outParts = New Collection
    For i = 1 To otherParts.Count
        If Trim$(otherParts(i)) <> "" Then outParts.Add otherParts(i)
    Next i

    If saleDict.Count > 0 Then
        Dim keyList() As Double, k As Variant
        Dim j As Long, a As Long, b As Long, tmp As Double
        ReDim keyList(0 To saleDict.Count - 1)
        j = 0
        For Each k In saleDict.Keys
            keyList(j) = CDbl(k)
            j = j + 1
        Next k
        For a = LBound(keyList) To UBound(keyList) - 1
            For b = a + 1 To UBound(keyList)
                If keyList(b) < keyList(a) Then
                    tmp = keyList(a): keyList(a) = keyList(b): keyList(b) = tmp
                End If
            Next b
        Next a
        For a = LBound(keyList) To UBound(keyList)
            priceKey = Format$(keyList(a), "0.00")
            Dim g As Variant
            g = saleDict(priceKey)
            outParts.Add "Sold " & g(0) & " @Total=" & FormatCurrency(g(1)) & " Profit=" & FormatCurrency(g(2))
        Next a
    End If

    For i = 1 To outParts.Count
        If rebuilt <> "" Then rebuilt = rebuilt & " | "
        rebuilt = rebuilt & outParts(i)
    Next i

    AggregateSaleNotes = rebuilt
End Function

' =====================================================================================
' DASHBOARD
' =====================================================================================
Private Sub EnsureBusinessSelector(wsDash As Worksheet)
    wsDash.Range("A3").value = "Business:"
    wsDash.Range("A3").Font.Bold = True
    If Trim$(CStr(wsDash.Range("B3").value)) = "" Then wsDash.Range("B3").value = ActiveBusinessName()
    With wsDash.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=BIZ_ALL & "," & BIZ_VG & "," & BIZ_PK & "," & BIZ_CL
    End With
    SaveConfigValue CFG_ACTIVE_BIZ, NormalizeBusiness(CStr(wsDash.Range("B3").value))
End Sub

Public Sub BuildDashboard()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Dashboard")
    ws.Cells.Clear
    ApplySheetChrome ws, False

    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 18
    ws.Range("A1:H1").Merge
    ws.Range("A2:H2").Merge
    ws.Range("A1").value = BRAND_TITLE
    ws.Range("A2").value = BRAND_SUB

    ws.Range("A1:H2").Interior.Color = CLR_PRIMARY
    ws.Range("A1:H2").Font.Color = CLR_SECONDARY
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Font.Italic = True

    ws.Range("A6").value = "Total Items"
    ws.Range("A7").value = "Unique Platforms"
    ws.Range("A8").value = "Total Value"
    ws.Range("A9").value = "Low Stock Count"
    ws.Range("A6:A9").Font.Bold = True

    ws.Range("B6").value = 0
    ws.Range("B7").value = 0
    ws.Range("B8").value = 0
    ws.Range("B9").value = 0
    ws.Range("B8").NumberFormat = "$#,##0.00"

    EnsureBusinessSelector ws
    EnsureDashboardButtons
    RecolorButtons ws
End Sub

Public Sub EnsureDashboardButtons()
    Dim ws As Worksheet, shp As Shape, hasBtn As Boolean
    Set ws = EnsureSheet("Dashboard")

    hasBtn = False
    For Each shp In ws.Shapes
        If shp.Name = "btnDash_Inventory" Then hasBtn = True: Exit For
    Next
    If hasBtn Then Exit Sub

    Dim colOffset As Double, rowOffset As Double
    colOffset = ws.Columns("D").left
    rowOffset = ws.Rows(4).top + 5

    AddAppButton ws, "Inventory", "GoInventory", colOffset, rowOffset, 135, 30, "btnDash_Inventory"
    AddAppButton ws, "Add Stock", "GoAddStockForm", colOffset + 150, rowOffset, 135, 30, "btnDash_AddStock"
    AddAppButton ws, "Batch Entry", "GoBatchEntry", colOffset + 300, rowOffset, 135, 30, "btnDash_Batch"
    AddAppButton ws, "Barcode Lookup", "GoLookup", colOffset + 450, rowOffset, 135, 30, "btnDash_Lookup"
    AddAppButton ws, "History", "GoHistory", colOffset + 600, rowOffset, 135, 30, "btnDash_History"
    AddAppButton ws, "Refresh", "RefreshDashboard", colOffset + 750, rowOffset, 135, 30, "btnDash_Refresh"
    AddAppButton ws, "Reset App Layout", "ResetAppLayoutsAndButtons", colOffset + 300, rowOffset, 135, 30, "btnDash_ResetAppLayout"
    rowOffset = rowOffset + 50
    AddAppButton ws, "Reset Settings", "ConfirmResetSettings", colOffset, rowOffset, 135, 30, "btnDash_ResetSettings"
    AddAppButton ws, "Toggle Theme", "ToggleDarkMode", colOffset + 150, rowOffset, 135, 30, "btnDash_ToggleTheme"
End Sub

Public Sub RefreshDashboard(): SafeRun "Refresh Dashboard", "RefreshDashboard_Worker": End Sub

Private Function UniqueCountFiltered(ws As Worksheet, ByVal colIdx As Long, ByRef cols As InventoryCols, ByVal bizFilter As String) As Long
    Dim d As Object, r As Long, lastRow As Long, v As String
    If colIdx = 0 Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    lastRow = TableLastRow(ws, "A")
    For r = 2 To lastRow
        If BusinessMatchesFilter(GetRowBusiness(ws, r, cols, True), bizFilter) Then
            v = Trim$(CStr(ws.Cells(r, colIdx).value))
            If Len(v) > 0 Then If Not d.Exists(v) Then d.Add v, 1
        End If
    Next r
    UniqueCountFiltered = d.Count
End Function

Private Sub RefreshDashboard_Worker()
    On Error GoTo ErrHandler

    Dim wsDash As Worksheet, wsInv As Worksheet, cols As InventoryCols
    Dim totalItems As Double, totalValue As Double, lowStock As Long
    Dim lastRow As Long, r As Long, lowThreshold As Long, bizFilter As String

    Set wsDash = EnsureSheet("Dashboard")
    Set wsInv = EnsureSheet("Inventory")
    EnsureInventoryHasHeadersAndBusiness
    GetInvCols wsInv, cols

    EnsureBusinessSelector wsDash
    bizFilter = ActiveBusinessName()

    lowThreshold = CLng(Val(GetConfigValue("LowStockThreshold", 5)))
    lastRow = TableLastRow(wsInv, "A")
    totalItems = 0: totalValue = 0: lowStock = 0

    For r = 2 To lastRow
        If BusinessMatchesFilter(GetRowBusiness(wsInv, r, cols, True), bizFilter) Then
            totalItems = totalItems + Val(wsInv.Cells(r, cols.Quantity).value)
            totalValue = totalValue + Val(wsInv.Cells(r, cols.Quantity).value) * Val(wsInv.Cells(r, cols.ListedPrice).value)
            If Val(wsInv.Cells(r, cols.Quantity).value) < lowThreshold Then lowStock = lowStock + 1
        End If
    Next r

    wsDash.Range("B6").value = totalItems
    wsDash.Range("B7").value = UniqueCountFiltered(wsInv, cols.Console, cols, bizFilter)
    wsDash.Range("B8").value = totalValue
    wsDash.Range("B9").value = lowStock
    wsDash.Range("B8").NumberFormat = "$#,##0.00"
    wsDash.Range("B8").Font.Color = IIf(totalValue > 0, CLR_SUCCESS, CLR_NEUTRAL)
    wsDash.Range("B9").Font.Color = IIf(lowStock > 0, CLR_ERROR, CLR_NEUTRAL)

    Exit Sub
ErrHandler:
    MsgBox "Error refreshing dashboard: " & Err.Number & " - " & Err.Description, vbExclamation, MSG_TITLE
End Sub

' =====================================================================================
' SETTINGS REPAIR / MIGRATION (FROM YOUR SNIPPETS)
' =====================================================================================
Public Sub FixSettingsPollution_RebuildPerBusinessLists()
    SafeRun "Fix Settings (Per-Business Dropdowns)", "FixSettingsPollution_RebuildPerBusinessLists_Worker"
End Sub

Private Sub FixSettingsPollution_RebuildPerBusinessLists_Worker()
    EnsureThemeReady

    Dim ws As Worksheet
    Set ws = EnsureSheet("Settings")

    ' 1) Capture the polluted list from column A (old "Business" column)
    Dim oldLast As Long, r As Long, v As String
    Dim items As Collection
    Set items = New Collection

    oldLast = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If oldLast < 2 Then oldLast = 2

    For r = 2 To oldLast
        v = Trim$(CStr(ws.Cells(r, "A").value))
        If v <> "" Then
            items.Add v
        End If
    Next r

    ' 2) Hard reset Settings sheet layout
    ws.Cells.Clear

    ' 3) Build clean per-business settings columns + named ranges
    BuildCleanSettingsSchema ws

    ' 4) Rebuild Business list (ONLY the 3 businesses)
    ws.Range("A1").value = "Business"
    ws.Range("A2").value = "Video Games"
    ws.Range("A3").value = "Pokemon"
    ws.Range("A4").value = "Cologne"

    ' 5) Move captured polluted items into Console / Product - Video Games (Column B)
    Dim outRow As Long: outRow = 2
    For r = 1 To items.Count
        v = Trim$(CStr(items(r)))
        If v <> "" Then
            If LCase$(v) <> "video games" And LCase$(v) <> "pokemon" And LCase$(v) <> "cologne" Then
                ws.Cells(outRow, "B").value = v
                outRow = outRow + 1
            End If
        End If
    Next r

    ' 6) Ensure defaults for Pokemon / Cologne if blank (starter values)
    If Trim$(CStr(ws.Cells(2, "C").value)) = "" Then ws.Cells(2, "C").value = "Card"
    If Trim$(CStr(ws.Cells(3, "C").value)) = "" Then ws.Cells(3, "C").value = "Sealed"
    If Trim$(CStr(ws.Cells(4, "C").value)) = "" Then ws.Cells(4, "C").value = "Graded"
    If Trim$(CStr(ws.Cells(2, "D").value)) = "" Then ws.Cells(2, "D").value = "Other"

    ' 7) Recreate named ranges used by validations
    RecreateBusinessAndPerBizNamedRanges ws

    ' 8) Cosmetic
    ApplySheetChrome ws, False
    NormalizeTopRows ws
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = CLR_HEADER
    ws.Columns.AutoFit

    MsgBox "Settings fixed:" & vbCrLf & _
           "- Business list reset to Video Games / Pokemon / Cologne" & vbCrLf & _
           "- Polluted items moved into Console / Product - Video Games list" & vbCrLf & _
           "- Per-business dropdown columns and named ranges rebuilt", vbInformation, MSG_TITLE
End Sub

Private Sub BuildCleanSettingsSchema(ws As Worksheet)
    ws.Range("A1").value = "Business"

    ws.Range("B1").value = "Console / Product - Video Games"
    ws.Range("C1").value = "Console / Product - Pokemon"
    ws.Range("D1").value = "Console / Product - Cologne"

    ws.Range("E1").value = "Quality - Video Games"
    ws.Range("F1").value = "Quality - Pokemon"
    ws.Range("G1").value = "Quality - Cologne"

    ws.Range("H1").value = "From - Video Games"
    ws.Range("I1").value = "From - Pokemon"
    ws.Range("J1").value = "From - Cologne"

    ws.Range("K1").value = "Listed Status - Video Games"
    ws.Range("L1").value = "Listed Status - Pokemon"
    ws.Range("M1").value = "Listed Status - Cologne"

    ws.Range("N1").value = "Stored - Video Games"
    ws.Range("O1").value = "Stored - Pokemon"
    ws.Range("P1").value = "Stored - Cologne"

    ws.Range("Q1").value = "Tested - Video Games"
    ws.Range("R1").value = "Tested - Pokemon"
    ws.Range("S1").value = "Tested - Cologne"

    ws.Range("T1").value = "Bundle Number - Video Games"
    ws.Range("U1").value = "Bundle Number - Pokemon"
    ws.Range("V1").value = "Bundle Number - Cologne"

    ws.Range("W1").value = "Country - Video Games"
    ws.Range("X1").value = "Country - Pokemon"
    ws.Range("Y1").value = "Country - Cologne"
End Sub

Private Sub RecreateBusinessAndPerBizNamedRanges(ws As Worksheet)
    ' Business selector list
    On Error Resume Next: ThisWorkbook.names("BusinessList").Delete: On Error GoTo 0
    ThisWorkbook.names.Add Name:="BusinessList", RefersTo:="='" & ws.Name & "'!" & ws.Range("A2:A4").Address

    ' Console / Product
    RecreateNamedRangeForColumn ws, "Field_console___product_Video_GamesList", "B"
    RecreateNamedRangeForColumn ws, "Field_console___product_PokemonList", "C"
    RecreateNamedRangeForColumn ws, "Field_console___product_CologneList", "D"

    ' Quality
    RecreateNamedRangeForColumn ws, "Field_quality_Video_GamesList", "E"
    RecreateNamedRangeForColumn ws, "Field_quality_PokemonList", "F"
    RecreateNamedRangeForColumn ws, "Field_quality_CologneList", "G"

    ' From
    RecreateNamedRangeForColumn ws, "Field_from_Video_GamesList", "H"
    RecreateNamedRangeForColumn ws, "Field_from_PokemonList", "I"
    RecreateNamedRangeForColumn ws, "Field_from_CologneList", "J"

    ' Listed Status
    RecreateNamedRangeForColumn ws, "Field_listed_status_Video_GamesList", "K"
    RecreateNamedRangeForColumn ws, "Field_listed_status_PokemonList", "L"
    RecreateNamedRangeForColumn ws, "Field_listed_status_CologneList", "M"

    ' Stored
    RecreateNamedRangeForColumn ws, "Field_stored_Video_GamesList", "N"
    RecreateNamedRangeForColumn ws, "Field_stored_PokemonList", "O"
    RecreateNamedRangeForColumn ws, "Field_stored_CologneList", "P"

    ' Tested
    RecreateNamedRangeForColumn ws, "Field_tested_Video_GamesList", "Q"
    RecreateNamedRangeForColumn ws, "Field_tested_PokemonList", "R"
    RecreateNamedRangeForColumn ws, "Field_tested_CologneList", "S"

    ' Bundle Number
    RecreateNamedRangeForColumn ws, "Field_bundle_number_Video_GamesList", "T"
    RecreateNamedRangeForColumn ws, "Field_bundle_number_PokemonList", "U"
    RecreateNamedRangeForColumn ws, "Field_bundle_number_CologneList", "V"

    ' Country
    RecreateNamedRangeForColumn ws, "Field_country_Video_GamesList", "W"
    RecreateNamedRangeForColumn ws, "Field_country_PokemonList", "X"
    RecreateNamedRangeForColumn ws, "Field_country_CologneList", "Y"

    ' Also refresh the "Console_*" and "Quality_*" named ranges used by the app:
    EnsureConsoleNamedRanges
    EnsureQualityNamedRanges
End Sub

Private Sub RecreateNamedRangeForColumn(ws As Worksheet, ByVal rangeName As String, ByVal colLetter As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    On Error Resume Next
    ThisWorkbook.names(rangeName).Delete
    On Error GoTo 0

    ThisWorkbook.names.Add Name:=rangeName, RefersTo:="='" & ws.Name & "'!" & ws.Range(colLetter & "2:" & colLetter & lastRow).Address
End Sub

' =====================================================================================
' VISUAL REPAIR / THEME REAPPLY
' =====================================================================================
Public Sub RepairInventoryColors()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        RecolorButtons ws
    Next
    RefreshDashboard_Worker
    On Error GoTo 0
End Sub

Private Sub ReapplyThemeAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ApplySheetChrome ws, False
        RecolorButtons ws
    Next
    EnsureDashboardButtons
    SetupStockHistoryButtons
    RefreshDashboard_Worker
End Sub

' =====================================================================================
' NAVIGATION
' =====================================================================================
Public Sub GoInventory(): EnsureSheet("Inventory").Activate: End Sub
Public Sub GoAddStockForm(): EnsureSheet("Add Stock").Activate: End Sub
Public Sub GoLookup(): EnsureSheet("Barcode Lookup").Activate: End Sub
Public Sub GoHistory(): EnsureSheet("Stock History").Activate: End Sub
Public Sub GoBatchEntry(): EnsureSheet("Batch Entry").Activate: End Sub

' =====================================================================================
' RESET SETTINGS
' =====================================================================================
Public Sub ConfirmResetSettings()
    If MsgBox("Reset settings to defaults?", vbYesNo + vbQuestion, MSG_TITLE) = vbYes Then ResetSettings
End Sub

Public Sub ResetSettings()
    SetupConfig
    SaveConfigValue CFG_ACTIVE_BIZ, BIZ_ALL
    MsgBox "Settings reset (Config).", vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' FULL APP REBUILD (LAYOUTS + BUTTONS)
' =====================================================================================
Public Sub ResetAppLayoutsAndButtons()
    SafeRun "Reset App Layouts + Buttons", "ResetAppLayoutsAndButtons_Worker"
End Sub

Private Sub ResetAppLayoutsAndButtons_Worker()
    EnsureThemeReady

    ' 1) Ensure core sheets exist / config is initialized
    SetupConfig
    EnsureSettingsBasics
    EnsureInventoryHasHeadersAndBusiness
    SetupBarcodesSheet False

    ' 2) Rebuild layouts
    BuildDashboard
    SetupAddStockForm
    SetupBatchEntry
    SetupBarcodeLookup
    SetupStockHistoryButtons

    ' 3) Rebind dropdowns everywhere (ensures Console/Product and Quality are linked)
    RebindAddStockDropdowns
    RebindBatchDropdowns
    RebindInventoryDropdowns

    ' 4) Final visual polish
    RepairInventoryColors
    RefreshDashboard_Worker

    MsgBox "Rebuilt layouts and rebound dropdowns on all pages.", vbInformation, MSG_TITLE
End Sub

' =====================================================================================
' BOOTSTRAP
' =====================================================================================
Public Sub SetupMultiBusinessSupport()
    SetupConfig
    EnsureSettingsBasics
    EnsureInventoryHasHeadersAndBusiness
    SetupBarcodesSheet False
    SetupAddStockForm
    SetupBatchEntry
    SetupBarcodeLookup
    BuildDashboard
    SetupStockHistoryButtons
    RebindInventoryDropdowns
    RefreshDashboard_Worker
    MsgBox "Initialized multi-business inventory app (VG / Pokemon / Cologne).", vbInformation, MSG_TITLE
End Sub

' Backward compatible name
Public Sub SetupTwoBusinessSupport()
    SetupMultiBusinessSupport
End Sub

' =========================
' END
' =========================
' =========================
' Settings (per-business named ranges) - FIXED to use Field_* names everywhere
' =========================

Private Sub EnsureSettingsBasics()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Settings")
    ApplySheetChrome ws, False

    ' Ensure schema exists (headers)
    If Trim$(CStr(ws.Range("A1").value)) = "" Then
        BuildCleanSettingsSchema ws
    Else
        ' If someone has old layout, still ensure correct headers exist
        BuildCleanSettingsSchema ws
    End If

    ' Business list
    ws.Range("A1").value = "Business"
    ws.Range("A2").value = BIZ_VG
    ws.Range("A3").value = BIZ_PK
    ws.Range("A4").value = BIZ_CL

    ' Create/refresh ALL Field_* named ranges
    RecreateBusinessAndPerBizNamedRanges ws

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = CLR_HEADER
End Sub

' Returns the correct Field_* list formula for a field + business
Private Function FieldListFormula(ByVal fieldKey As String, ByVal bizVal As String) As String
    bizVal = NormalizeBusiness(bizVal)
    Select Case bizVal
        Case BIZ_PK: FieldListFormula = "=Field_" & fieldKey & "_PokemonList"
        Case BIZ_CL: FieldListFormula = "=Field_" & fieldKey & "_CologneList"
        Case Else:   FieldListFormula = "=Field_" & fieldKey & "_Video_GamesList"
    End Select
End Function

' Convenience wrappers
Private Function BusinessConsoleNamedList(ByVal bizVal As String) As String
    ' console___product (triple underscore is intentional: "console___product")
    BusinessConsoleNamedList = FieldListFormula("console___product", bizVal)
End Function

Private Function BusinessQualityNamedList(ByVal bizVal As String) As String
    BusinessQualityNamedList = FieldListFormula("quality", bizVal)
End Function

Private Function BusinessFromNamedList(ByVal bizVal As String) As String
    BusinessFromNamedList = FieldListFormula("from", bizVal)
End Function

Private Function BusinessListedStatusNamedList(ByVal bizVal As String) As String
    BusinessListedStatusNamedList = FieldListFormula("listed_status", bizVal)
End Function

Private Function BusinessStoredNamedList(ByVal bizVal As String) As String
    BusinessStoredNamedList = FieldListFormula("stored", bizVal)
End Function

Private Function BusinessTestedNamedList(ByVal bizVal As String) As String
    BusinessTestedNamedList = FieldListFormula("tested", bizVal)
End Function

Private Function BusinessBundleNamedList(ByVal bizVal As String) As String
    BusinessBundleNamedList = FieldListFormula("bundle_number", bizVal)
End Function

Private Function BusinessCountryNamedList(ByVal bizVal As String) As String
    BusinessCountryNamedList = FieldListFormula("country", bizVal)
End Function

' =========================
' Add Stock form dropdown binding (FIXED)
' =========================

Private Sub BindAddStockDropdowns_All(ws As Worksheet)
    EnsureSettingsBasics
    EnsureAddStockBusinessField ws

    BindAddStockConsoleDropdown ws
    BindAddStockQualityDropdown ws
    BindAddStockFromDropdown ws
    BindAddStockListedStatusDropdown ws
    BindAddStockStoredDropdown ws
    BindAddStockTestedDropdown ws
    BindAddStockBundleDropdown ws
    BindAddStockCountryDropdown ws
End Sub

Private Sub BindAddStockConsoleDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim bizCell As Range, tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessConsoleNamedList(bizVal)
    Set tgt = AddStockField(ws, "Console / Product")
    If tgt Is Nothing Then Set tgt = AddStockField(ws, "Console")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockQualityDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessQualityNamedList(bizVal)
    Set tgt = AddStockField(ws, "Quality")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockFromDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessFromNamedList(bizVal)
    Set tgt = AddStockField(ws, "From")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockListedStatusDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessListedStatusNamedList(bizVal)
    Set tgt = AddStockField(ws, "Listed Status")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockStoredDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessStoredNamedList(bizVal)
    Set tgt = AddStockField(ws, "Stored")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockTestedDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessTestedNamedList(bizVal)
    Set tgt = AddStockField(ws, "Tested")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockBundleDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessBundleNamedList(bizVal)
    Set tgt = AddStockField(ws, "Bundle Number")
    If tgt Is Nothing Then Set tgt = AddStockField(ws, "Bundle")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Private Sub BindAddStockCountryDropdown(ws As Worksheet)
    Dim bizVal As String, formulaToUse As String
    Dim tgt As Range

    bizVal = NormalizeBusiness(FieldVal(ws, "Business"))
    If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG

    formulaToUse = BusinessCountryNamedList(bizVal)
    Set tgt = AddStockField(ws, "Country")
    If tgt Is Nothing Then Exit Sub

    With tgt.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
    End With
End Sub

Public Sub RebindAddStockDropdowns()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Add Stock")
    BindAddStockDropdowns_All ws
    MsgBox "Add Stock dropdowns rebound to current Business.", vbInformation, MSG_TITLE
End Sub

' Make sure SetupAddStockForm calls the new binder:
' In your SetupAddStockForm, replace:
'   EnsureAddStockBusinessField ws
'   BindAddStockConsoleDropdown ws
'   BindAddStockQualityDropdown ws
' with:
'   BindAddStockDropdowns_All ws

' =========================
' Batch Entry dropdown binding (FIXED)
' =========================

Private Sub BindBatchDropdowns_All(ws As Worksheet)
    EnsureSettingsBasics
    ApplyBatchBusinessDropdown ws
    BindBatchConsoleDropdowns ws
    BindBatchQualityDropdowns ws
End Sub

Private Sub BindBatchConsoleDropdowns(ws As Worksheet)
    Dim bizCol As Long, conCol As Long, r As Long
    Dim bizVal As String, formulaToUse As String

    bizCol = MatchCol(ws, "Business")
    conCol = MatchConsoleColAny(ws)
    If bizCol = 0 Or conCol = 0 Then Exit Sub

    For r = 2 To 300
        bizVal = NormalizeBusiness(CStr(ws.Cells(r, bizCol).value))
        If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG
        formulaToUse = BusinessConsoleNamedList(bizVal)

        With ws.Cells(r, conCol).Validation
            .Delete
            .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
        End With
    Next r
End Sub

Private Sub BindBatchQualityDropdowns(ws As Worksheet)
    Dim bizCol As Long, qualCol As Long, r As Long
    Dim bizVal As String, formulaToUse As String

    bizCol = MatchCol(ws, "Business")
    qualCol = MatchCol(ws, "Quality")
    If bizCol = 0 Or qualCol = 0 Then Exit Sub

    For r = 2 To 300
        bizVal = NormalizeBusiness(CStr(ws.Cells(r, bizCol).value))
        If bizVal = "" Or bizVal = BIZ_ALL Then bizVal = BIZ_VG
        formulaToUse = BusinessQualityNamedList(bizVal)

        With ws.Cells(r, qualCol).Validation
            .Delete
            .Add xlValidateList, xlValidAlertStop, xlBetween, formulaToUse
        End With
    Next r
End Sub

Public Sub RebindBatchDropdowns()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Batch Entry")
    BindBatchDropdowns_All ws
    MsgBox "Batch Entry dropdowns rebound based on Business column values.", vbInformation, MSG_TITLE
End Sub

' Make sure SetupBatchEntry calls the new binder:
' After headers/layout, replace:
'   ApplyBatchBusinessDropdown ws
'   BindBatchConsoleDropdowns ws
'   BindBatchQualityDropdowns ws
' with:
'   BindBatchDropdowns_All ws
