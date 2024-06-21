这个PowerShell脚本可以令您基于选定的既有dotx模板，来新建一个Word文档。

请务必**在目标文件夹**打开PowerShell。

````
# 加载Windows.Forms程序集
Add-Type -AssemblyName System.Windows.Forms

# 创建文件选择对话框函数
function Show-FileDialog {
    param (
        [string]$title,
        [string]$filter
    )
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = $filter
    $fileDialog.Title = $title
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    } else {
        return $null
    }
}

# 选择.dotx模板文件
$templatePath = Show-FileDialog -title "请选择.dotx模板文件" -filter "Word模板文件 (*.dotx)|*.dotx"
if (-not $templatePath) {
    Write-Host "未选择模板文件，脚本结束"
    exit
}

# 获取当前目录
$targetFolder = (Get-Location).Path

# 加载Word应用程序
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

try {
    # 创建新文档
    Write-Host "使用模板创建新文档..."
    $document = $wordApp.Documents.Add($templatePath)
    $newDocumentPath = [System.IO.Path]::Combine($targetFolder, "New Document.docx")
    Write-Host "保存新文档到路径：$newDocumentPath"
    $document.SaveAs([ref] $newDocumentPath, [ref] [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument)
    $document.Close()
    Write-Host "新文档已成功创建，保存路径为：$newDocumentPath"
} catch {
    Write-Host "创建新文档时出错：$_"
} finally {
    # 释放Word应用程序对象
    Write-Host "释放Word应用程序对象..."
    $wordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp)
}

Write-Host "脚本运行结束。"
````
