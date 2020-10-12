$currentPath = Split-Path -Parent $PSCommandPath

#���W���[���̓ǂݍ���
$itexDLL = Join-Path $currentPath "itextsharp.dll"
[Reflection.Assembly]::LoadFrom($itexDLL) | Out-Null

function ReadPdfFile($_filePath){
    if(Test-Path -LiteralPath $_filePath){
        $reader = New-Object iTextSharp.text.pdf.PdfReader($_filePath)
        #Pdf�̍ő�y�[�W���擾
        $pages = $reader.NumberOfPages
        for($page = 1; $page -le $pages; $page++){
            #1�y�[�W���ǂݍ���
            $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page)
            #1�s���ɕ�����
            $lines = $text -split "\n"
            foreach($line in $lines){
                Write-Host $line
            }
        }

        $reader.Close()
        $reader.Dispose()
    }else{
        Write-Host $_filePath ��������܂���ł����B
    }
}

#����ΏۂƂ��Ă���Pdf�t�@�C�����ǂ������m�F����
function GetTargetExtension($_fileExtension){
    $result = $false
    #Pdf�t�@�C���݂̂�ΏۂƂ���
    Select-String -InputObject $_fileExtension -Pattern ".pdf" | ForEach-Object { $_.Matches } | ForEach-Object { $result = $true }
    return $result
}

#�����Ƃ��Ď󂯎�����p�X����g���q���擾����
function GetFileName($_filePath){
    $fileExtension = [System.IO.Path]::GetExtension($_filePath)
    $result = GetTargetExtension -_fileExtension $fileExtension
    if($result){
        ReadPdfFile -_filePath $_filePath
    }else {
        Write-Host �ΏۂƂȂ�t�@�C���`���ł͂���܂���ł����B
    }
}

$filePath = "Pdf�t�@�C���܂ł̃p�X"
GetFileName -_filePath $filePath