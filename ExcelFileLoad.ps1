#excelファイルを読み込む
function ReadExcelFile($_filePath){
    #ファイルの存在を確認する
    if(Test-Path $_filePath){
        $excel = New-Object -ComObject Excel.Application
        try{
            #excelを非表示で実行
            $excel.Visible = $false
            #警告ウィンドウを表示しない
            $excel.DisplayAlerts = $false
            #エクセルを開く
            $book =  $excel.Workbooks.Open($_filePath)
            #シートを読み込み
            $book.Worksheets | ForEach-Object {
                #ハイパーリンク取得
                $_.Hyperlinks | ForEach-Object {
                    Write-Host $_.Address
                }

                #図形にあるテキストボックスから取得
                $_.Shapes | ForEach-Object {
                    #グループの場合(msoGroupの値は6)
                    if($_.Type -eq 6){
                        $_.GroupItems | ForEach-Object{
                            if($_.TextFrame2.HasText){
                                Write-Host $_.TextFrame2.TextRange.Text
                            }
                        }
                    }else{
                        if($_.TextFrame2.HasText){
                            Write-Host $_.TextFrame2.TextRange.Text
                        }
                    }
                }
                
                #行読み込み
                $_.UsedRange.Rows | ForEach-Object {
                    #列を読み込み
                    $_.Columns | ForEach-Object {
                        #文字が記載されている個所のみ取得
                        if($_.Text -ne ""){
                            Write-Host $_.Text
                        }
                    }
                }
            }
        } catch {
            Write-Host $_.Exception.ToString()
        } finally {
            #excelファイル操作終了時の決まった処理(はじまり)
            if($book){
                $book.Close()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
                $book = $null
                Remove-Variable -Name book -ErrorAction SilentlyContinue
            }

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            if($excel){
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                $excel = $null
                Remove-Variable -Name excel -ErrorAction SilentlyContinue
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()
            }
            #excelファイル操作終了時の決まった処理(おわり)
        }
    }else{
        Write-Host $_filePath + " が見つかりませんでした。"
    }
}

#今回対象としているExcelファイルかどうかを確認する
function GetTargetExtension($_fileName){
    $result = $false
    #excelファイルのみを対象とする
    Select-String -InputObject $fileName -Pattern ".+\.(xlsx?|xlsm)" | ForEach-Object { $_.Matches } | ForEach-Object { $result = $true }
    return $result
}

#引数として受け取ったパスからファイル名(拡張子含む)を取得する
function GetFileName($_filePath){
    $fileName = [System.IO.Path]::GetFileName($_filePath)
    $result = GetTargetExtension -_fileName $fileName
    if($result){
        ReadExcelFile -_filePath $_filePath
    }else {
        Write-Host 対象となるファイル形式ではありませんでした。
    }
}

#xls, xlsx, xlsmは動作確認済み
$filePath = ""
GetFileName -_filePath $filePath