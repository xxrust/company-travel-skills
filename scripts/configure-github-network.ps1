[CmdletBinding(DefaultParameterSetName = 'Configure')]
param(
    [Parameter()]
    [string]$Proxy = 'http://127.0.0.1:7890',

    [Parameter()]
    [switch]$Persist,

    [Parameter()]
    [switch]$Clear,

    [Parameter()]
    [switch]$SkipGit,

    [Parameter()]
    [switch]$SkipNpm,

    [Parameter()]
    [switch]$SkipTest
)

$ErrorActionPreference = 'Stop'

function Set-ProcessProxy([string]$Value) {
    # These variables are inherited by Git, npm, npx and other child processes.
    $env:HTTP_PROXY = $Value
    $env:HTTPS_PROXY = $Value
    $env:ALL_PROXY = $Value
}

function Clear-ProcessProxy {
    Remove-Item Env:HTTP_PROXY -ErrorAction SilentlyContinue
    Remove-Item Env:HTTPS_PROXY -ErrorAction SilentlyContinue
    Remove-Item Env:ALL_PROXY -ErrorAction SilentlyContinue
}

function Set-UserProxy([string]$Value) {
    # Windows PowerShell can block while broadcasting an environment change
    # through .NET on some machines. Writing HKCU\Environment is immediate;
    # new processes will inherit these values.
    foreach ($name in @('HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY')) {
        & reg.exe ADD 'HKCU\Environment' /v $name /t REG_SZ /d $Value /f | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "写入用户环境变量 $name 失败"
        }
    }
}

function Clear-UserProxy {
    foreach ($name in @('HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY')) {
        & reg.exe DELETE 'HKCU\Environment' /v $name /f 2>$null | Out-Null
    }
}

if ($Clear) {
    Clear-ProcessProxy
    Clear-UserProxy

    if (-not $SkipGit) {
        git config --global --unset http.proxy 2>$null
        git config --global --unset https.proxy 2>$null
    }

    if (-not $SkipNpm -and (Get-Command npm -ErrorAction SilentlyContinue)) {
        npm config delete proxy --location=user
        npm config delete https-proxy --location=user
    }

    Write-Host '已清除当前用户的显式代理环境变量、Git 全局代理和 npm 用户代理配置。'
    exit 0
}

try {
    $proxyUri = [Uri]$Proxy
    if (-not $proxyUri.IsAbsoluteUri -or $proxyUri.Scheme -notin @('http', 'https', 'socks5', 'socks5h')) {
        throw "代理地址必须是 http、https、socks5 或 socks5h URL：$Proxy"
    }
}
catch {
    throw "无效的代理地址：$Proxy。$($_.Exception.Message)"
}

Set-ProcessProxy $Proxy
if ($Persist) {
    Set-UserProxy $Proxy
}

if (-not $SkipGit) {
    git config --global http.proxy $Proxy
    git config --global https.proxy $Proxy
}

if (-not $SkipNpm -and (Get-Command npm -ErrorAction SilentlyContinue)) {
    npm config set proxy $Proxy --location=user
    npm config set https-proxy $Proxy --location=user
}

Write-Host "当前 PowerShell 进程已设置 HTTP_PROXY、HTTPS_PROXY、ALL_PROXY=$Proxy"
if ($Persist) {
    Write-Host '已写入当前用户环境变量；新开的 PowerShell 窗口会自动继承。'
}
if (-not $SkipGit) {
    Write-Host '已设置 Git 全局 HTTP/HTTPS 代理。'
}
if (-not $SkipNpm -and (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Host '已设置 npm 用户代理，npx 会继承该配置。'
}

if (-not $SkipTest) {
    Write-Host '正在测试 GitHub HTTPS 访问...'
    $response = Invoke-WebRequest -Uri 'https://api.github.com' -UseBasicParsing -TimeoutSec 20
    if ($response.StatusCode -ne 200) {
        throw "GitHub API 返回 HTTP $($response.StatusCode)"
    }

    $probe = git ls-remote https://github.com/vercel-labs/agent-skills.git HEAD 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Git 访问 GitHub 失败：$probe"
    }

    Write-Host 'GitHub API 和 Git HTTPS 访问测试通过。'
}
