' VBA_Optimization.bas - Optimisation des macros VBA pour une exécution plus rapide
' Exemple : Comparaison avant/après optimisation

Attribute VB_Name = "VBA_Optimization"

Sub OptimizedMacro()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim arr As Variant
    Dim i As Long
    
    ' Désactiver les mises à jour pour accélérer l'exécution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Définir la feuille et la plage de données
    Set ws = ActiveSheet
    Set dataRange = ws.Range("A2:A10000") ' Exemple : 10 000 lignes
    
    ' Charger les données dans un tableau VBA (Variant)
    arr = dataRange.Value
    
    ' Traitement des données en mémoire
    For i = LBound(arr) To UBound(arr)
        arr(i, 1) = arr(i, 1) * 2 ' Exemple : multiplier chaque valeur par 2
    Next i
    
    ' Réinjecter les données dans la feuille en une seule opération
    dataRange.Value = arr
    
    ' Réactiver les mises à jour
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Optimisation terminée !", vbInformation, "Succès"
End Sub
