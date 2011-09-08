# 事前に PowerShell で
# Set-ExecutionPolicy RemoteSigned CurrentUser
# を実行しておき、署名なしでもスクリプトを実行できるようにしておく。
# パスの通ったところに less.exe を置いておく。
# <MYDOCUMENT>\WindowsPowerShell\Modules\MyModule\MyModule.psm1
# にこのファイルを置く。
# <MYDOCUMENT>\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
# に ipmo MyModule の行を追加する。

Set-StrictMode -Version 2.0

# Pager
sal more.com less.exe

<#
.SYNOPSIS
Get or set window width
.PARAMETER w

.EXAMPLE
width
Get window width

.EXAMPLE
width 100
Set window width to 100

.EXAMPLE
width ((width)+10)
Add window width by 10

.LINK
height

.NOTES
mode.com
[Console]::WindowWidth, [Console]::BufferWidth
#>
function width([int]$w=0){
	if($w -le 0) {
		return $Host.UI.RawUI.WindowSize.Width;
	}
	$ws = $Host.UI.RawUI.WindowSize;
	$bs = $Host.UI.RawUI.BufferSize;
	$ws.Width = $w;
	$bs.Width = $w;
	if ( $w -gt $Host.UI.RawUI.BufferSize.Width) {
		$Host.UI.RawUI.BufferSize = $bs;
		$Host.UI.RawUI.WindowSize = $ws;
	} else {
		$Host.UI.RawUI.WindowSize = $ws;
		$Host.UI.RawUI.BufferSize = $bs;
	}
}
<#
.SYNOPSIS
Get or set window height
.PARAMETER h

.EXAMPLE
height
Get window height

.EXAMPLE
height 40
Set window height to 40

.EXAMPLE
height ((height)+5)
Add window height by 5

.LINK
width

.NOTES
mode.com
[Console]::WindowHeight, [Console]::BufferHeight
#>
function height([int]$h=0){
	if($h -le 0) {
		return $Host.UI.RawUI.WindowSize.Height;
	}
	$ws = $Host.UI.RawUI.WindowSize;
	$ws.Height = $h;
	$Host.UI.RawUI.WindowSize = $ws;
}

<#
.SYNOPSIS
Get or set window title
.PARAMETER t

.EXAMPLE
title
Get window title

.EXAMPLE
title hoge
Set window title to hoge

.NOTES
[Console]::Title
#>
function title([string]$t="") {
	if($t -eq "") {
		return $Host.UI.RawUI.WindowTitle;
	}
	$Host.UI.RawUI.WindowTitle = $t;
}

Export-ModuleMember -Alias * -Function * -Cmdlet *
