#EDINETAPI色々取得ver0.1
#URL
class EdinetDocs {
    $url = "https://disclosure.edinet-fsa.go.jp/api/v1/documents.json"
    [Object[]]$docs
    [int]$sum

    EdinetDocs([String]$targetDate) {
        $request = $this.url + "?date=" + $targetDate + "&type=2"
        $res = Invoke-RestMethod $request
        #連続実行の可能性を加味してWait
        Start-Sleep -Seconds 3

        if ($res.metadata.status -ne "200") {
            throw "ステータスコード " + $res.metadata.status + " が返却されました"
        }

        $this.sum = $res.metadata.resultset.count
        $this.docs = $res.results
    }

    [Object[]]showDocs(){
        return $this.docs
    }
}

class Filter {
    [EdinetDocs]$edinetDocs
    [String]$filerNameFilter=""
    [String]$docDescriptionFilter=""

    Filter([EdinetDocs]$docs) {
        $this.edinetDocs = $docs
    }

    [Void]setFilerNameFilter([String]$arg){
        $this.filerNameFilter = $arg
    }

    [Void]setDocDescriptionFilter([String]$arg){
        $this.docDescriptionFilter = $arg
    }

    hidden [Object[]]apply(){
        $let = $this.edinetDocs.showDocs() `
                | Where-Object -Property 'docDescription' -Match $this.docDescriptionFilter `
                | Where-Object -Property 'filerName' -Match $this.filerNameFilter `
                | Select-Object -Property 'docID','docDescription','filerName' 
        return $let
    }

    [Object[]]showByCui(){
        return $this.apply()
    }

    [Object[]]showByGridView(){
        return $this.apply() | Out-GridView -PassThru -Title "EDINET取得結果"
    }

    [Void]showByCsv(){
        $here = Get-Location
        $now = Get-Date -Format FileDateTime 
        $fileName = Join-Path -Path $here -ChildPath "$now.csv"
        $this.apply() | Export-Csv -Path $fileName -Encoding unicode
    }
}

#how to use getMetaDataFromEdinet
#$date:=日付 ex)"yyyy-MM-DD" [必須項目]
#$docsName:=ドキュメント名の検索条件
#$filerName:=提出者（会社）名の検索条件
#$outputPattern:="Cui","GridView","Csv" [必須項目]
#ex) getDocsFromEdinet -date "2022-08-19" -outputPattern Cui -filerName "三菱" -docsName "半期"
function getMetaDataFromEdinet() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$date,
        [String]$docsName,
        [String]$filerName,
        [Parameter(Mandatory)]
        [ValidateSet("Cui","GridView","Csv")]
        [String]$outputPattern
    )

    #new
    if ($date -match '^[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]$') {
        $docs = [EdinetDocs]::new($date)
        $filter = [Filter]::new($docs)
    }else{
        throw "ただしい形式の日付を入力して下さい。ex)yyyy-MM-DD"
    }

    #set filter
    if ($docsName.Length -gt 0 ) {
        $filter.setDocDescriptionFilter($docsName)
    }

    if ($filerName.Length -gt 0 ) {
        $filter.setFilerNameFilter($filerName)
    }

    #Output
    if ($outputPattern -eq "Cui") {
        $filter.showByCui()
    }

    if ($outputPattern -eq "GridView") {
        $filter.showByGridView()
    }

    if ($outputPattern -eq "Csv") {
        $filter.showByCsv()
    }
}

#how to use getDocZipFromEdinet
#$docID を渡す。
function getDocZipFromEdinet([String]$docID) {
    $Url="https://disclosure.edinet-fsa.go.jp/api/v1/documents/" + $docID + "?type=1"
    $dir = Join-Path -Path (Get-Location) -ChildPath "zips"
    $zipName = Join-Path -Path $dir -ChildPath "$docID.zip"
    
    if ((Test-Path $dir) -eq $false) {
        mkdir -Path $dir
    }

    Invoke-RestMethod  $Url -OutFile $zipName

    #連続の実行を考慮してwait
    Start-Sleep -Seconds 3
}