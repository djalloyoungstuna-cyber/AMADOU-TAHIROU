# 📊 Système de Gestion des Ressources Humaines (SGRH)

## Vue d'ensemble
Ce document fournit un système complet de gestion des ressources humaines avec des procédures stockées sécurisées, des scripts VBA et des commandes PowerShell.

---

## 1. 🗄️ Code SQL (Procédures Stockées)

### 1.1 Création de la Table Employés

```sql
CREATE TABLE Employees (
    EmployeeID INT PRIMARY KEY IDENTITY(1,1),
    FirstName NVARCHAR(100) NOT NULL,
    LastName NVARCHAR(100) NOT NULL,
    Email NVARCHAR(100) UNIQUE NOT NULL,
    PhoneNumber NVARCHAR(20),
    HireDate DATETIME NOT NULL,
    Department NVARCHAR(100) NOT NULL,
    Position NVARCHAR(100) NOT NULL,
    Salary DECIMAL(12, 2) NOT NULL,
    ManagerID INT,
    Status NVARCHAR(20) DEFAULT 'Active',
    CreatedAt DATETIME DEFAULT GETDATE(),
    UpdatedAt DATETIME DEFAULT GETDATE(),
    FOREIGN KEY (ManagerID) REFERENCES Employees(EmployeeID)
);
```

### 1.2 Procédure Stockée - Ajouter un Employé

```sql
CREATE PROCEDURE dbo.AddEmployee
    @FirstName NVARCHAR(100),
    @LastName NVARCHAR(100),
    @Email NVARCHAR(100),
    @PhoneNumber NVARCHAR(20),
    @HireDate DATETIME,
    @Department NVARCHAR(100),
    @Position NVARCHAR(100),
    @Salary DECIMAL(12, 2),
    @ManagerID INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Validation des paramètres
    IF @FirstName IS NULL OR LEN(TRIM(@FirstName)) = 0
    BEGIN
        RAISERROR('Le prénom est obligatoire', 16, 1);
        RETURN;
    END
    
    IF @LastName IS NULL OR LEN(TRIM(@LastName)) = 0
    BEGIN
        RAISERROR('Le nom de famille est obligatoire', 16, 1);
        RETURN;
    END
    
    IF @Email IS NULL OR LEN(TRIM(@Email)) = 0
    BEGIN
        RAISERROR('L''email est obligatoire', 16, 1);
        RETURN;
    END
    
    IF @Salary < 0
    BEGIN
        RAISERROR('Le salaire ne peut pas être négatif', 16, 1);
        RETURN;
    END
    
    -- Vérification de l'email unique
    IF EXISTS (SELECT 1 FROM Employees WHERE Email = @Email)
    BEGIN
        RAISERROR('Cet email existe déjà dans le système', 16, 1);
        RETURN;
    END
    
    -- Insertion de l'employé
    INSERT INTO Employees (FirstName, LastName, Email, PhoneNumber, HireDate, Department, Position, Salary, ManagerID, Status)
    VALUES (@FirstName, @LastName, @Email, @PhoneNumber, @HireDate, @Department, @Position, @Salary, @ManagerID, 'Active');
    
    SELECT SCOPE_IDENTITY() AS EmployeeID;
END;
```

### 1.3 Procédure Stockée - Mettre à Jour un Employé

```sql
CREATE PROCEDURE dbo.UpdateEmployee
    @EmployeeID INT,
    @FirstName NVARCHAR(100) = NULL,
    @LastName NVARCHAR(100) = NULL,
    @Email NVARCHAR(100) = NULL,
    @PhoneNumber NVARCHAR(20) = NULL,
    @Department NVARCHAR(100) = NULL,
    @Position NVARCHAR(100) = NULL,
    @Salary DECIMAL(12, 2) = NULL,
    @ManagerID INT = NULL,
    @Status NVARCHAR(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Vérifier que l'employé existe
    IF NOT EXISTS (SELECT 1 FROM Employees WHERE EmployeeID = @EmployeeID)
    BEGIN
        RAISERROR('Employé non trouvé', 16, 1);
        RETURN;
    END
    
    -- Mettre à jour les champs fournis
    UPDATE Employees
    SET 
        FirstName = COALESCE(@FirstName, FirstName),
        LastName = COALESCE(@LastName, LastName),
        Email = COALESCE(@Email, Email),
        PhoneNumber = COALESCE(@PhoneNumber, PhoneNumber),
        Department = COALESCE(@Department, Department),
        Position = COALESCE(@Position, Position),
        Salary = COALESCE(@Salary, Salary),
        ManagerID = COALESCE(@ManagerID, ManagerID),
        Status = COALESCE(@Status, Status),
        UpdatedAt = GETDATE()
    WHERE EmployeeID = @EmployeeID;
    
    SELECT 'Employé mis à jour avec succès' AS Message;
END;
```

### 1.4 Procédure Stockée - Récupérer les Employés par Département

```sql
CREATE PROCEDURE dbo.GetEmployeesByDepartment
    @Department NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        EmployeeID,
        FirstName,
        LastName,
        Email,
        PhoneNumber,
        HireDate,
        Department,
        Position,
        Salary,
        Status,
        CreatedAt
    FROM Employees
    WHERE Department = @Department AND Status = 'Active'
    ORDER BY FirstName, LastName;
END;
```

### 1.5 Procédure Stockée - Calculer la Masse Salariale par Département

```sql
CREATE PROCEDURE dbo.GetSalaryByDepartment
    @Department NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        Department,
        COUNT(*) AS NumberOfEmployees,
        SUM(Salary) AS TotalSalary,
        AVG(Salary) AS AverageSalary,
        MIN(Salary) AS MinimumSalary,
        MAX(Salary) AS MaximumSalary
    FROM Employees
    WHERE Status = 'Active' 
        AND (@Department IS NULL OR Department = @Department)
    GROUP BY Department
    ORDER BY Department;
END;
```

### 1.6 Procédure Stockée - Supprimer (Désactiver) un Employé

```sql
CREATE PROCEDURE dbo.DeactivateEmployee
    @EmployeeID INT,
    @Reason NVARCHAR(500)
AS
BEGIN
    SET NOCOUNT ON;
    
    IF NOT EXISTS (SELECT 1 FROM Employees WHERE EmployeeID = @EmployeeID)
    BEGIN
        RAISERROR('Employé non trouvé', 16, 1);
        RETURN;
    END
    
    UPDATE Employees
    SET 
        Status = 'Inactive',
        UpdatedAt = GETDATE()
    WHERE EmployeeID = @EmployeeID;
    
    -- Enregistrer dans l'audit (optionnel)
    INSERT INTO EmployeeAudit (EmployeeID, Action, Reason, ActionDate)
    VALUES (@EmployeeID, 'Deactivation', @Reason, GETDATE());
    
    SELECT 'Employé désactivé avec succès' AS Message;
END;
```

---

## 2. 📝 Code VBA (Application Client)

### 2.1 Module VBA - Ajouter un Employé

```vba
Option Explicit

Sub AddNewEmployee()
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strConn As String
    Dim EmployeeID As Long
    
    ' Chaîne de connexion SQL Server
    strConn = "Provider=SQLOLEDB;Data Source=YourServerName;Initial Catalog=YourDatabaseName;User ID=YourUsername;Password=YourPassword;"
    
    ' Initialisation de la connexion
    Set conn = New ADODB.Connection
    conn.Open strConn
    
    ' Initialisation de la commande
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.AddEmployee"
    
    ' Ajout des paramètres
    cmd.Parameters.Append cmd.CreateParameter("@FirstName", adVarWChar, adParamInput, 100, "Jean")
    cmd.Parameters.Append cmd.CreateParameter("@LastName", adVarWChar, adParamInput, 100, "Dupont")
    cmd.Parameters.Append cmd.CreateParameter("@Email", adVarWChar, adParamInput, 100, "jean.dupont@company.com")
    cmd.Parameters.Append cmd.CreateParameter("@PhoneNumber", adVarWChar, adParamInput, 20, "06 12 34 56 78")
    cmd.Parameters.Append cmd.CreateParameter("@HireDate", adDBDate, adParamInput, , Date())
    cmd.Parameters.Append cmd.CreateParameter("@Department", adVarWChar, adParamInput, 100, "IT")
    cmd.Parameters.Append cmd.CreateParameter("@Position", adVarWChar, adParamInput, 100, "Développeur")
    cmd.Parameters.Append cmd.CreateParameter("@Salary", adCurrency, adParamInput, , 35000)
    cmd.Parameters.Append cmd.CreateParameter("@ManagerID", adInteger, adParamInput, , 1)
    
    ' Exécution de la commande
    On Error GoTo ErrorHandler
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        EmployeeID = rs.Fields("EmployeeID").Value
        MsgBox "Employé ajouté avec succès. ID: " & EmployeeID, vbInformation, "Succès"
    Else
        MsgBox "Erreur lors de l'ajout de l'employé.", vbCritical, "Erreur"
    End If
    
    ' Fermeture des objets
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, "Erreur"
    If Not conn Is Nothing Then conn.Close
End Sub
```

### 2.2 Module VBA - Afficher les Employés par Département

```vba
Sub DisplayEmployeesByDepartment()
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strConn As String
    Dim ws As Worksheet
    Dim rowIndex As Long
    
    ' Chaîne de connexion
    strConn = "Provider=SQLOLEDB;Data Source=YourServerName;Initial Catalog=YourDatabaseName;User ID=YourUsername;Password=YourPassword;"
    
    ' Initialisation de la feuille
    Set ws = ThisWorkbook.Sheets("Employees")
    ws.Cells.Clear
    
    ' En-têtes
    ws.Range("A1:H1").Value = Array("ID", "Prénom", "Nom", "Email", "Téléphone", "Département", "Poste", "Salaire")
    
    ' Connexion
    Set conn = New ADODB.Connection
    conn.Open strConn
    
    ' Commande
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.GetEmployeesByDepartment"
    cmd.Parameters.Append cmd.CreateParameter("@Department", adVarWChar, adParamInput, 100, "IT")
    
    ' Exécution
    On Error GoTo ErrorHandler
    Set rs = cmd.Execute
    
    ' Remplir le tableau
    rowIndex = 2
    Do While Not rs.EOF
        ws.Cells(rowIndex, 1).Value = rs.Fields("EmployeeID").Value
        ws.Cells(rowIndex, 2).Value = rs.Fields("FirstName").Value
        ws.Cells(rowIndex, 3).Value = rs.Fields("LastName").Value
        ws.Cells(rowIndex, 4).Value = rs.Fields("Email").Value
        ws.Cells(rowIndex, 5).Value = rs.Fields("PhoneNumber").Value
        ws.Cells(rowIndex, 6).Value = rs.Fields("Department").Value
        ws.Cells(rowIndex, 7).Value = rs.Fields("Position").Value
        ws.Cells(rowIndex, 8).Value = rs.Fields("Salary").Value
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    MsgBox "Données affichées avec succès", vbInformation, "Succès"
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, "Erreur"
End Sub
```

---

## 3. 🖥️ Scripts PowerShell

### 3.1 Script PowerShell - Gérer les Employés

```powershell
# Paramètres de connexion
param(
    [string]$ServerName = "YourServerName",
    [string]$DatabaseName = "YourDatabaseName",
    [string]$Username = "YourUsername",
    [string]$Password = "YourPassword",
    [string]$Action = "Add"
)

# Fonction pour ajouter un employé
function Add-Employee {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Email,
        [string]$PhoneNumber,
        [string]$Department,
        [string]$Position,
        [decimal]$Salary,
        [int]$ManagerID = $null
    )
    
    $ConnectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$Username;Password=$Password;"
    
    $SqlCommand = @"
    DECLARE @EmployeeID INT;
    EXEC dbo.AddEmployee
        @FirstName = '$FirstName',
        @LastName = '$LastName',
        @Email = '$Email',
        @PhoneNumber = '$PhoneNumber',
        @HireDate = GETDATE(),
        @Department = '$Department',
        @Position = '$Position',
        @Salary = $Salary,
        @ManagerID = $ManagerID;
    SELECT @EmployeeID AS EmployeeID;
"@
    
    try {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()
        $Command = New-Object System.Data.SqlClient.SqlCommand($SqlCommand, $Connection)
        $Reader = $Command.ExecuteReader()
        if ($Reader.Read()) {
            $EmployeeID = $Reader["EmployeeID"]
            Write-Host "✓ Employé ajouté avec succès. ID: $EmployeeID" -ForegroundColor Green
        } else {
            Write-Host "✗ Erreur lors de l'ajout de l'employé." -ForegroundColor Red
        }
        $Reader.Close()
        $Connection.Close()
    } catch {
        Write-Host "✗ Erreur: $_" -ForegroundColor Red
    }
}

# Fonction pour récupérer les employés par département
function Get-EmployeesByDepartment {
    param([string]$Department)
    
    $ConnectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$Username;Password=$Password;"
    
    $SqlCommand = @"
    EXEC dbo.GetEmployeesByDepartment @Department = '$Department';
"@
    
    try {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()
        $Command = New-Object System.Data.SqlClient.SqlCommand($SqlCommand, $Connection)
        $Reader = $Command.ExecuteReader()
        
        $employees = @()
        while ($Reader.Read()) {
            $employees += @{
                EmployeeID = $Reader["EmployeeID"]
                FirstName = $Reader["FirstName"]
                LastName = $Reader["LastName"]
                Email = $Reader["Email"]
                Position = $Reader["Position"]
                Salary = $Reader["Salary"]
            }
        }
        
        $Reader.Close()
        $Connection.Close()
        
        return $employees
    } catch {
        Write-Host "✗ Erreur: $_" -ForegroundColor Red
    }
}

# Fonction pour obtenir les statistiques salariales
function Get-SalaryStatistics {
    param([string]$Department = $null)
    
    $ConnectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$Username;Password=$Password;"
    
    $deptFilter = if ($Department) { "'$Department'" } else { "NULL" }
    
    $SqlCommand = @"
    EXEC dbo.GetSalaryByDepartment @Department = $deptFilter;
"@
    
    try {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()
        $Command = New-Object System.Data.SqlClient.SqlCommand($SqlCommand, $Connection)
        $Reader = $Command.ExecuteReader()
        
        Write-Host "`n=== STATISTIQUES SALARIALES ===" -ForegroundColor Cyan
        Write-Host ""
        
        while ($Reader.Read()) {
            Write-Host "Département: $($Reader['Department'])" -ForegroundColor Yellow
            Write-Host "  Nombre d'employés: $($Reader['NumberOfEmployees'])"
            Write-Host "  Masse salariale totale: $($Reader['TotalSalary']:C)"
            Write-Host "  Salaire moyen: $($Reader['AverageSalary']:C)"
            Write-Host "  Salaire minimum: $($Reader['MinimumSalary']:C)"
            Write-Host "  Salaire maximum: $($Reader['MaximumSalary']:C)"
            Write-Host ""
        }
        
        $Reader.Close()
        $Connection.Close()
    } catch {
        Write-Host "✗ Erreur: $_" -ForegroundColor Red
    }
}

# Exécution des fonctions
switch ($Action) {
    "Add" {
        Add-Employee -FirstName "Marie" -LastName "Martin" -Email "marie.martin@company.com" `
                     -PhoneNumber "06 98 76 54 32" -Department "HR" -Position "Responsable RH" -Salary 40000
    }
    "GetByDept" {
        Get-EmployeesByDepartment -Department "IT" | Format-Table
    }
    "Salary" {
        Get-SalaryStatistics
    }
    default {
        Write-Host "Action non reconnue. Utilisez: Add, GetByDept, ou Salary" -ForegroundColor Red
    }
}
```

---

## 4. 🔒 Sécurité & Bonnes Pratiques

### 4.1 Recommandations de Sécurité

1. **Validation des Données**
   - Validez tous les inputs (longueur, format, type)
   - Utilisez toujours des paramètres pour éviter les injections SQL
   - Limitez les caractères spéciaux

2. **Chiffrement**
   - Chiffrez les mots de passe dans les chaînes de connexion
   - Utilisez `SecureString` en PowerShell
   - Utilisez SSL/TLS pour les communications

3. **Audit et Logs**
   - Enregistrez toutes les modifications (AddEmployee, UpdateEmployee, DeactivateEmployee)
   - Tracez les accès et les modifications par utilisateur
   - Conservez l'historique pendant 7 ans

4. **Contrôle d'Accès**
   - Limitez les permissions aux rôles appropriés
   - Utilisez le principe du moindre privilège
   - Implémentez l'authentification multi-facteurs

5. **Sauvegardes**
   - Effectuez des sauvegardes quotidiennes
   - Testez régulièrement la restauration
   - Archivez les sauvegardes hors site

---

## 5. 📋 Table d'Audit (Optionnel)

```sql
CREATE TABLE EmployeeAudit (
    AuditID INT PRIMARY KEY IDENTITY(1,1),
    EmployeeID INT NOT NULL,
    Action NVARCHAR(50),
    OldValue NVARCHAR(MAX),
    NewValue NVARCHAR(MAX),
    Reason NVARCHAR(500),
    ActionDate DATETIME DEFAULT GETDATE(),
    ActionBy NVARCHAR(100) DEFAULT SUSER_NAME(),
    FOREIGN KEY (EmployeeID) REFERENCES Employees(EmployeeID)
);
```

---

## 6. 🚀 Déploiement & Exécution

### 6.1 Étapes d'Installation

1. **Créer la base de données**
   ```sql
   CREATE DATABASE HumanResources;
   USE HumanResources;
   ```

2. **Exécuter les scripts SQL**
   - Créez la table `Employees`
   - Créez toutes les procédures stockées
   - Créez la table d'audit

3. **Configurer les connexions**
   - Mettez à jour les chaînes de connexion dans VBA et PowerShell
   - Testez la connexion à la base de données

4. **Importer les données initiales**
   ```sql
   EXEC dbo.AddEmployee 'Admin', 'System', 'admin@company.com', '01 00 00 00 00', GETDATE(), 'IT', 'Administrator', 50000, NULL
   ```

---

## 7. 📊 Rapports Disponibles

- Employés par département
- Masse salariale par département
- Historique des modifications
- Rapports de conformité
- Analyse de l'ancienneté

---

## 8. 📞 Support & Documentation

Pour toute question ou assistance supplémentaire, consultez la documentation SQL Server ou contactez l'équipe IT.

**Dernière mise à jour**: 2026-06-13
**Version**: 1.0
