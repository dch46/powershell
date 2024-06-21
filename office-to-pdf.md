下面的PowerShell脚本可以调用Office 365的相关功能，将Office文件批量转换为PDF文件。仅在Office 365桌面版起作用。

````
# 加载Windows.Forms和Drawing程序集
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 创建文件选择对话框函数
function Show-FileDialog {
    param (
        [string]$title,
        [string]$filter
    )
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = $filter
    $fileDialog.Title = $title
    $fileDialog.Multiselect = $true
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileNames
    } else {
        return $null
    }
}

# 创建文件夹选择对话框函数
function Show-FolderDialog {
    param (
        [string]$description
    )
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = $description
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderDialog.SelectedPath
    } else {
        return $null
    }
}

# 定义转换函数：将PowerPoint文件转换为PDF
function Convert-PowerPointToPDF {
    param (
        [string]$inputFile,
        [string]$outputFolder
    )
    try {
        $outputFile = [System.IO.Path]::ChangeExtension("$outputFolder\$([System.IO.Path]::GetFileNameWithoutExtension($inputFile))", "pdf")
        $presentation = $pptApp.Presentations.Open($inputFile, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        $presentation.SaveAs($outputFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
        $presentation.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
    } catch {
        Write-Host "Error converting $inputFile to PDF: $_"
        if ($presentation) {
            $presentation.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
        }
    }
}

# 定义转换函数：将Word文件转换为PDF
function Convert-WordToPDF {
    param (
        [string]$inputFile,
        [string]$outputFolder
    )
    try {
        $outputFile = [System.IO.Path]::ChangeExtension("$outputFolder\$([System.IO.Path]::GetFileNameWithoutExtension($inputFile))", "pdf")
        $document = $wordApp.Documents.Open($inputFile)
        $document.SaveAs([ref] $outputFile, [ref] [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF)
        $document.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document)
    } catch {
        Write-Host "Error converting $inputFile to PDF: $_"
        if ($document) {
            $document.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document)
        }
    }
}

# 定义转换函数：将Excel文件转换为PDF
function Convert-ExcelToPDF {
    param (
        [string]$inputFile,
        [string]$outputFolder
    )
    try {
        $outputFile = [System.IO.Path]::ChangeExtension("$outputFolder\$([System.IO.Path]::GetFileNameWithoutExtension($inputFile))", "pdf")
        $workbook = $excelApp.Workbooks.Open($inputFile)
        $workbook.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $outputFile)
        $workbook.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    } catch {
        Write-Host "Error converting $inputFile to PDF: $_"
        if ($workbook) {
            $workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
        }
    }
}

# 选择文件
$filePaths = Show-FileDialog -title "请选择需要转换的文件" -filter "所有支持的文件 (*.ppt; *.pptx; *.doc; *.docx; *.xls; *.xlsx)|*.ppt;*.pptx;*.doc;*.docx;*.xls;*.xlsx"
if (-not $filePaths) {
    Write-Host "未选择文件，脚本结束"
    exit
}

# 提示用户选择输出路径选项
Write-Host "请选择输出路径选项："
Write-Host "A - 原文件路径"
Write-Host "B - 在原文件路径下新建 'Converted' 文件夹"
Write-Host "C - 指定路径"
$outputOption = Read-Host "请输入 A, B 或 C"

# 根据用户选择的输出路径选项设置输出文件夹
switch ($outputOption.ToUpper()) {
    "A" {
        $outputFolder = [System.IO.Path]::GetDirectoryName($filePaths[0])
    }
    "B" {
        $outputFolder = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($filePaths[0]), "Converted")
        if (-not (Test-Path -Path $outputFolder)) {
            New-Item -ItemType Directory -Path $outputFolder
        }
    }
    "C" {
        $outputFolder = Show-FolderDialog -description "请选择输出PDF文件的文件夹"
        if (-not $outputFolder) {
            Write-Host "未选择输出文件夹，脚本结束"
            exit
        }
    }
    default {
        Write-Host "未知选项，脚本结束"
        exit
    }
}

# 遍历选中的文件并转换
foreach ($filePath in $filePaths) {
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

    switch ($extension) {
        ".ppt" {
            # 加载PowerPoint应用程序
            if (-not $pptApp) {
                $pptApp = New-Object -ComObject PowerPoint.Application
                $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
            }
            Convert-PowerPointToPDF $filePath $outputFolder
        }
        ".pptx" {
            # 加载PowerPoint应用程序
            if (-not $pptApp) {
                $pptApp = New-Object -ComObject PowerPoint.Application
                $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
            }
            Convert-PowerPointToPDF $filePath $outputFolder
        }
        ".doc" {
            # 加载Word应用程序
            if (-not $wordApp) {
                $wordApp = New-Object -ComObject Word.Application
                $wordApp.Visible = $false
            }
            Convert-WordToPDF $filePath $outputFolder
        }
        ".docx" {
            # 加载Word应用程序
            if (-not $wordApp) {
                $wordApp = New-Object -ComObject Word.Application
                $wordApp.Visible = $false
            }
            Convert-WordToPDF $filePath $outputFolder
        }
        ".xls" {
            # 加载Excel应用程序
            if (-not $excelApp) {
                $excelApp = New-Object -ComObject Excel.Application
                $excelApp.Visible = $false
            }
            Convert-ExcelToPDF $filePath $outputFolder
        }
        ".xlsx" {
            # 加载Excel应用程序
            if (-not $excelApp) {
                $excelApp = New-Object -ComObject Excel.Application
                $excelApp.Visible = $false
            }
            Convert-ExcelToPDF $filePath $outputFolder
        }
        default {
            Write-Host "不支持的文件类型：$extension"
        }
    }
}

# 释放应用程序对象
if ($pptApp) {
    $pptApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp)
}

if ($wordApp) {
    $wordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp)
}

if ($excelApp) {
    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
}

Write-Host "文件转换完成，PDF保存在：$outputFolder"
````
