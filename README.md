### **1. Code SQL (Procédure Stockée Sécurisée)**
#### **Exemple de procédure stockée (SQL Server)**
```sql
CREATE PROCEDURE dbo.InsertOrder
    @CustomerID INT,
    @OrderDate DATETIME = GETDATE(),
    @TotalAmount DECIMAL(10, 2)
AS
BEGIN
    SET NOCOUNT ON;

    -- Vérification des paramètres
    IF @CustomerID IS NULL OR @TotalAmount <= 0
    BEGIN
        RAISERROR('Paramètres invalides', 16, 1);
        RETURN;
    END

    -- Insertion de la commande
    INSERT INTO Orders (CustomerID, OrderDate, TotalAmount)
    VALUES (@CustomerID, @OrderDate, @TotalAmount);

    -- Récupération de l'ID de la commande insérée
    SELECT SCOPE_IDENTITY() AS OrderID;
END;
```

#### **Exemple de procédure stockée (Microsoft Access)**
```sql
CREATE PROCEDURE InsertOrder
    CustomerID INTEGER,
    OrderDate DATETIME,
    TotalAmount CURRENCY
AS
BEGIN
    -- Vérification des paramètres
    IF ISNULL(CustomerID) OR TotalAmount <= 0 THEN
        MsgBox "Paramètres invalides", vbCritical, "Erreur"
        EXIT PROCEDURE
    END IF

    -- Insertion de la commande
    INSERT INTO Orders (CustomerID, OrderDate, TotalAmount)
    VALUES (CustomerID, OrderDate, TotalAmount);

    -- Récupération de l'ID de la commande insérée
    SELECT @@IDENTITY AS OrderID;
END;
```

---

### **2. Code VBA (Application Client)**
#### **Exemple de code VBA pour Microsoft Access**
```vba
Option Explicit

Sub InsertOrder()
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim OrderID As Long

    ' Initialisation de la connexion
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\To\Your\Database.accdb;"

    ' Initialisation de la commande
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "InsertOrder"

    ' Ajout des paramètres
    cmd.Parameters.Append cmd.CreateParameter("CustomerID", adInteger, adParamInput, , 1) ' Remplacez 1 par l'ID du client
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBDate, adParamInput, , Date())
    cmd.Parameters.Append cmd.CreateParameter("TotalAmount", adCurrency, adParamInput, , 99.99) ' Remplacez 99.99 par le montant total

    ' Exécution de la commande
    Set rs = cmd.Execute
    If Not rs.EOF Then
        OrderID = rs.Fields("OrderID").Value
        MsgBox "Commande insérée avec succès. ID: " & OrderID, vbInformation, "Succès"
    Else
        MsgBox "Erreur lors de l'insertion de la commande.", vbCritical, "Erreur"
    End If

    ' Fermeture des objets
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub
```

#### **Exemple de code VBA pour SQL Server**
```vba
Option Explicit

Sub InsertOrder()
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strConn As String
    Dim OrderID As Long

    ' Chaîne de connexion SQL Server
    strConn = "Provider=SQLOLEDB;Data Source=YourServerName;Initial Catalog=YourDatabaseName;User ID=YourUsername;Password=YourPassword;"

    ' Initialisation de la connexion
    Set conn = New ADODB.Connection
    conn.Open strConn

    ' Initialisation de la commande
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.InsertOrder"

    ' Ajout des paramètres
    cmd.Parameters.Append cmd.CreateParameter("@CustomerID", adInteger, adParamInput, , 1) ' Remplacez 1 par l'ID du client
    cmd.Parameters.Append cmd.CreateParameter("@OrderDate", adDBDate, adParamInput, , Date())
    cmd.Parameters.Append cmd.CreateParameter("@TotalAmount", adCurrency, adParamInput, , 99.99) ' Remplacez 99.99 par le montant total

    ' Exécution de la commande
    Set rs = cmd.Execute
    If Not rs.EOF Then
        OrderID = rs.Fields("OrderID").Value
        MsgBox "Commande insérée avec succès. ID: " & OrderID, vbInformation, "Succès"
    Else
        MsgBox "Erreur lors de l'insertion de la commande.", vbCritical, "Erreur"
    End If

    ' Fermeture des objets
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub
```

---

### **3. Ligne de Commande Appropriée (PowerShell)**
#### **Exemple de script PowerShell pour exécuter la procédure stockée**
```powershell
# Paramètres
$ServerName = "YourServerName"
$DatabaseName = "YourDatabaseName"
$Username = "YourUsername"
$Password = "YourPassword"
$CustomerID = 1
$TotalAmount = 99.99

# Chaîne de connexion
$ConnectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$Username;Password=$Password;"

# Commande SQL
$SqlCommand = @"
DECLARE @OrderID INT;
EXEC dbo.InsertOrder
    @CustomerID = $CustomerID,
    @OrderDate = GETDATE(),
    @TotalAmount = $TotalAmount,
    @OrderID = @OrderID OUTPUT;
SELECT @OrderID AS OrderID;
"@

# Exécution de la commande
try {
    $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $Connection.Open()
    $Command = New-Object System.Data.SqlClient.SqlCommand($SqlCommand, $Connection)
    $Reader = $Command.ExecuteReader()
    if ($Reader.Read()) {
        $OrderID = $Reader["OrderID"]
        Write-Host "Commande insérée avec succès. ID: $OrderID"
    } else {
        Write-Host "Erreur lors de l'insertion de la commande."
    }
    $Reader.Close()
    $Connection.Close()
} catch {
    Write-Host "Erreur: $_"
} finally {
    if ($Connection.State -eq "Open") {
        $Connection.Close()
    }
}
```

---

### **4. Sécurité & Bonnes Pratiques**
1. **Paramètres de procédure stockée** :
   - Utilisez toujours des **paramètres** pour éviter les injections SQL.
   - Validez les entrées dans la procédure stockée.

2. **VBA** :
   - Utilisez **ADODB** pour les connexions et commandes.
   - Fermez toujours les objets **Connection**, **Command**, et **Recordset**.

3. **PowerShell** :
   - Utilisez des **variables** pour les informations sensibles (ex : mot de passe).
   - Gestion des erreurs avec **try/catch**.

4. **Chiffrement** :
   - Chiffrez les **mots de passe** dans les chaînes de connexion (ex : avec **SecureString** en PowerShell).

---

### **5. Exécution du Code**
1. **VBA** :
   - Ouvrez **Microsoft Access** ou **Excel**.
   - Allez dans **Outils > Éditeur VBA**.
   - Copiez-collez le code dans un module.
   - Exécutez la procédure **InsertOrder**.

2. **PowerShell** :
   - Ouvrez **PowerShell ISE** ou **VS Code**.
   - Copiez-collez le script.
   - Exécutez le script avec **F5**.

---

### **6. Exemple de Sortie**
#### **Sortie VBA (Microsoft Access)**
```
Commande insérée avec succès. ID: 123
```

#### **Sortie PowerShell**
```
Commande insérée avec succès. ID: 123
```

---

### **Prochaines Étapes**
1. **Ajouter des logs** pour le suivi des commandes.
2. **Intégrer des tests unitaires** pour vérifier la sécurité et la fiabilité.
3. **Déployer en production** avec des **procédures de sauvegarde** et de **restauration**.

