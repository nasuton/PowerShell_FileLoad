﻿#powerpointファイルを読み込む
function ReadPowerPointFile($_filePath){
    if(Test-Path $_filePath){
        $pptx = New-Object -ComObject PowerPoint.Application
        try {
            #非表示にしようとするとエラーがでる
            #$pptx.Visible = "msoFalse"
            #警告ウィンドウを表示しない
            #$pptx.DisplayAlerts = "ppAlertsNone"
            $slides = $pptx.presentations.Open($_filePath)
            $slides.Slides | ForEach-Object {
                #ハイパーリンク取得
                $_.Hyperlinks | ForEach-Object {
                    Write-Host $_.Address
                }

                #図形にあるテキストボックスから取得
                $_.Shapes | ForEach-Object {
                    #テキストボックス
                    if($_.HasTextFrame){
                        Write-Host $_.TextFrame.TextRange.text
                    }
                    #表
                    elseif($_.HasTable){
                        $_.Table.Columns | ForEach-Object {
                            $_ | ForEach-Object {
                                Write-Host $_.Shape.TextFrame.TextRange.text
                            }
                        }
                    #グラフ
                    }elseif($_.HasChart){
                        if($_.Chart.HasTitle){
                            Write-Host $_.Chart.Title
                        }
                    }
                    #スマートアート
                    elseif($_.HasSmartArt){
                        $_.SmartArt.Nodes | ForEach-Object {
                            Write-Host $_.TextFrame2.TextRange.text
                        }
                    }
                    #グループの場合(msoGroupの値は6)
                    elseif($_.Type -eq 6){
                        $shp.GroupItems | ForEach-Object {
                            if($_.HasTextFrame){
                                Write-Host $_.TextFrame.TextRange.text
                            }
                        }
                    }
                }
            }
        } catch {
            Write-Host $_.Exception.ToString()
        } finally {
            #powerpointファイル操作終了時の決まった処理(はじまり)
            if($slides){
                $slides.Close()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($slides) | Out-Null
                $slides = $null
                Remove-Variable -Name slides -ErrorAction SilentlyContinue
            }

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            if($pptx){
                $pptx.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptx) | Out-Null
                $pptx = $null
                Remove-Variable -Name pptx -ErrorAction SilentlyContinue
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()
            }
            #powerpointファイル操作終了時の決まった処理(おわり)
        }
    }else{
        Write-Host $_filePath が見つかりませんでした。
    }
}

#今回対象としているPowerPointファイルかどうかを確認する
function GetTargetExtension($_fileName){
    $result = $false
    #excelファイルのみを対象とする
    Select-String -InputObject $fileName -Pattern ".+\.(pptx?|pptm)" | ForEach-Object { $_.Matches } | ForEach-Object { $result = $true }
    return $result
}

#引数として受け取ったパスからファイル名(拡張子含む)を取得する
function GetFileName($_filePath){
    $fileName = [System.IO.Path]::GetFileName($_filePath)
    $result = GetTargetExtension -_fileName $fileName
    if($result){
        ReadPowerPointFile -_filePath $_filePath
    }else {
        Write-Host 対象となるファイル形式ではありませんでした。
    }
}

#ppt, pptx, pptmは動作確認済み
$filePath = ""
GetFileName -_filePath $filePath