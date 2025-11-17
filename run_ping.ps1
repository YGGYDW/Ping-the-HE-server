# Ping HE 的IPV6隧道服务器
param()

# 可修改此处次数：1 或 10 或其它
$Count = 5

# 隧道服务器列表
$servers = @(
    @{Country='美国 弗吉尼亚州 阿什本';       IP='216.66.22.2';        Region='北美'},
    @{Country='加拿大 艾伯塔省 卡尔加里';     IP='216.218.200.58';     Region='北美'},
    @{Country='美国 伊利诺伊州 芝加哥';       IP='184.105.253.14';     Region='北美'},
    @{Country='美国 德克萨斯州 达拉斯';       IP='184.105.253.10';     Region='北美'},
    @{Country='美国 科罗拉多州 丹佛';         IP='184.105.250.46';     Region='北美'},
    @{Country='美国 加州 弗里蒙特';           IP='72.52.104.74';       Region='北美'},
    @{Country='美国 加州 弗里蒙特';           IP='64.62.134.130';      Region='北美'},
    @{Country='美国 夏威夷 檀香山';           IP='64.71.156.86';       Region='北美'},
    @{Country='美国 密苏里 堪萨斯城';         IP='216.66.77.230';      Region='北美'},
    @{Country='美国 洛杉矶';                   IP='66.220.18.42';       Region='北美'},
    @{Country='美国 佛州 迈阿密';             IP='209.51.161.58';      Region='北美'},
    @{Country='美国 纽约';                     IP='209.51.161.14';      Region='北美'},
    @{Country='美国 亚利桑那 凤凰城';         IP='66.220.7.82';        Region='北美'},
    @{Country='美国 华盛顿州 西雅图';         IP='216.218.226.238';    Region='北美'},
    @{Country='加拿大 多伦多';                 IP='216.66.38.58';       Region='北美'},
    @{Country='加拿大 温尼伯';                 IP='184.105.255.26';     Region='北美'},

    @{Country='德国 柏林';                     IP='216.66.86.114';      Region='欧洲'},
    @{Country='匈牙利 布达佩斯';               IP='216.66.87.14';       Region='欧洲'},
    @{Country='德国 法兰克福';                 IP='216.66.80.30';       Region='欧洲'},
    @{Country='葡萄牙 里斯本';                 IP='216.66.87.102';      Region='欧洲'},
    @{Country='英国 伦敦';              		 IP='216.66.80.26';       Region='欧洲'},
    @{Country='英国 伦敦';               		 IP='216.66.88.98';       Region='欧洲'},
    @{Country='法国 巴黎';                     IP='216.66.84.42';       Region='欧洲'},
    @{Country='捷克 布拉格';                   IP='216.66.86.122';      Region='欧洲'},
    @{Country='瑞典 斯德哥尔摩';               IP='216.66.80.90';       Region='欧洲'},
    @{Country='波兰 华沙';                     IP='216.66.80.162';      Region='欧洲'},
    @{Country='瑞士 苏黎世';                   IP='216.66.80.98';       Region='欧洲'},

    @{Country='香港';                           IP='216.218.221.6';      Region='亚洲'},
    @{Country='新加坡';                         IP='216.218.221.42';     Region='亚洲'},
    @{Country='日本 东京';                       IP='74.82.46.6';         Region='亚洲'},

    @{Country='吉布提';                         IP='216.66.87.98';       Region='非洲'},
    @{Country='南非 约翰内斯堡';               IP='216.66.87.134';      Region='非洲'},

    @{Country='哥伦比亚 波哥大';               IP='216.66.64.154';      Region='南美'},

    @{Country='澳大利亚 悉尼';                 IP='216.218.142.50';     Region='大洋洲'},

    @{Country='阿联酋 迪拜';                   IP='216.66.90.30';       Region='中东'}
)

# 选择导出格式（只保存其中之一）
function Choose-ExportFormat {
    while ($true) {
        Write-Host ""
        Write-Host "请选择导出格式："
        Write-Host "  1) CSV (兼容性强，适合脚本/自动化)"
        Write-Host "  2) XLSX (需要本机安装 Excel，支持格式/多表)"
        Write-Host "  Q) 取消并退出"
        $choice = Read-Host "输入 1 / 2 / Q 并按回车"
        switch ($choice.ToUpper()) {
            '1' { return 'CSV' }
            '2' { return 'XLSX' }
            'Q' { return 'QUIT' }
            default { Write-Host "无效选择，请重试。" }
        }
    }
}

# 主循环：ping 并解析
$results = @()
$lastRegion = $null

foreach ($s in $servers) {
    if ($s.Region -ne $lastRegion) {
        Write-Output ""
        Write-Output "===== $($s.Region) ====="
        $lastRegion = $s.Region
    }

    Write-Output "Testing $($s.Country) -> $($s.IP) (Ping次数=$Count)"
    $pingRaw = ping -n $Count $s.IP | Out-String

    # 提取 time 值（英文 time= 与中文 时间=）
    $times = New-Object System.Collections.Generic.List[int]
    $timeRegex = '(?:time[=<]\s*(\d+)ms)|(?:时间[=<]\s*(\d+)ms)'
    $timeMatches = [regex]::Matches($pingRaw, $timeRegex)
    foreach ($m in $timeMatches) {
        $val = $m.Groups[1].Value
        if ([string]::IsNullOrEmpty($val)) { $val = $m.Groups[2].Value }
        if (-not [string]::IsNullOrEmpty($val)) { $times.Add([int]$val) }
    }

    $avg = -1; $min = -1; $max = -1

    if ($times.Count -gt 0) {
        $min = ($times | Measure-Object -Minimum).Minimum
        $max = ($times | Measure-Object -Maximum).Maximum
        $rawAvg = ($times | Measure-Object -Average).Average
        if ($null -ne $rawAvg) { $avg = [int][math]::Round($rawAvg) } else { $avg = -1 }
    } else {
        # 回退解析汇总行（英文/中文）
        $m1 = [regex]::Match($pingRaw, 'Minimum =\s*(\d+)ms,\s*Maximum =\s*(\d+)ms,\s*Average =\s*(\d+)ms')
        if ($m1.Success) {
            $min = [int]$m1.Groups[1].Value
            $max = [int]$m1.Groups[2].Value
            $avg = [int]$m1.Groups[3].Value
        } else {
            $m2 = [regex]::Match($pingRaw, '最小 =\s*(\d+)ms,\s*最大 =\s*(\d+)ms,\s*平均 =\s*(\d+)ms')
            if ($m2.Success) {
                $min = [int]$m2.Groups[1].Value
                $max = [int]$m2.Groups[2].Value
                $avg = [int]$m2.Groups[3].Value
            } else {
                $m3 = [regex]::Match($pingRaw, 'Average =\s*(\d+)ms')
                if ($m3.Success) { $avg = [int]$m3.Groups[1].Value }
                else {
                    $m4 = [regex]::Match($pingRaw, '平均 =\s*(\d+)ms')
                    if ($m4.Success) { $avg = [int]$m4.Groups[1].Value }
                }
            }
        }
    }

    # 丢包率解析（多种形式），否则用 replyCount 回退
    $loss = 100
    $lossMatch = [regex]::Match($pingRaw, '\((\d+)%\s*loss\)')
    if ($lossMatch.Success) { $loss = [int]$lossMatch.Groups[1].Value }
    else {
        $lossMatch2 = [regex]::Match($pingRaw, '丢失.*?\((\d+)%')
        if ($lossMatch2.Success) { $loss = [int]$lossMatch2.Groups[1].Value }
        else {
            $lostMatch3 = [regex]::Match($pingRaw, 'Lost = \s*\d+ \((\d+)%')
            if ($lostMatch3.Success) { $loss = [int]$lostMatch3.Groups[1].Value }
            else {
                $replyCount = $timeMatches.Count
                if ($Count -gt 0) {
                    $loss = [int]([math]::Round((($Count - $replyCount) / $Count) * 100))
                    if ($loss -lt 0) { $loss = 0 }
                    if ($loss -gt 100) { $loss = 100 }
                }
            }
        }
    }

    # TTL 提取（多语言）
    $ttl = -1
    $ttlMatch = [regex]::Match($pingRaw, 'TTL=(\d+)')
    if ($ttlMatch.Success) { $ttl = [int]$ttlMatch.Groups[1].Value }
    else {
        $ttlMatch2 = [regex]::Match($pingRaw, '生存时间[=:]\s*(\d+)')
        if ($ttlMatch2.Success) { $ttl = [int]$ttlMatch2.Groups[1].Value }
        else {
            $ttlMatch3 = [regex]::Match($pingRaw, 'ttl[=:]\s*(\d+)')
            if ($ttlMatch3.Success) { $ttl = [int]$ttlMatch3.Groups[1].Value }
        }
    }

    $obj = [PSCustomObject]@{
        Country = $s.Country
        IP = $s.IP
        AverageMs = $avg
        MinMs = $min
        MaxMs = $max
        LossPct = $loss
        TTL = $ttl
        Region = $s.Region
    }
    $results += $obj
}

# 排序并添加 Rank
$ordered = $results | Sort-Object @{Expression = { if ($_.AverageMs -lt 0) { 999999 } else { $_.AverageMs }}}, LossPct
$rank = 1
$final = foreach ($r in $ordered) {
    $r | Add-Member -NotePropertyName Rank -NotePropertyValue $rank -PassThru
    $rank++
}

# 选择导出格式
$fmt = Choose-ExportFormat
if ($fmt -eq 'QUIT') {
    Write-Output "已取消，未生成文件。"
    exit 0
}

$outCsv = Join-Path (Get-Location) 'results.csv'
$outXlsx = Join-Path (Get-Location) 'results.xlsx'

if ($fmt -eq 'CSV') {
    $final | Select-Object Rank, Region, Country, IP, AverageMs, MinMs, MaxMs, LossPct, TTL |
        Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
    Write-Output "CSV 已生成： $outCsv"
    exit 0
}

# XLSX 路径（尝试 COM，失败时询问回退为 CSV）
if ($fmt -eq 'XLSX') {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $headers = 'Rank','Region','Country','IP','AverageMs','MinMs','MaxMs','LossPct','TTL'
        for ($i=0; $i -lt $headers.Count; $i++) {
            $ws.Cells.Item(1, $i+1) = $headers[$i]
        }
        $row = 2
        foreach ($r in $final) {
            $ws.Cells.Item($row,1) = $r.Rank
            $ws.Cells.Item($row,2) = $r.Region
            $ws.Cells.Item($row,3) = $r.Country
            $ws.Cells.Item($row,4) = $r.IP
            $ws.Cells.Item($row,5) = $r.AverageMs
            $ws.Cells.Item($row,6) = $r.MinMs
            $ws.Cells.Item($row,7) = $r.MaxMs
            $ws.Cells.Item($row,8) = $r.LossPct
            $ws.Cells.Item($row,9) = $r.TTL
            $row++
        }
        $wb.SaveAs($outXlsx)
        $wb.Close($false)
        $excel.Quit()
        Write-Output "XLSX 已生成： $outXlsx"
        exit 0
    } catch {
        Write-Output "生成 XLSX 失败：未检测到可用 Excel COM 或发生错误。"
        $fallback = Read-Host "是否回退保存为 CSV？(Y/N)"
        if ($fallback.ToUpper() -eq 'Y') {
            $final | Select-Object Rank, Region, Country, IP, AverageMs, MinMs, MaxMs, LossPct, TTL |
                Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
            Write-Output "CSV 已生成： $outCsv"
            exit 0
        } else {
            Write-Output "已取消导出。"
            exit 1
        }
    }
}
