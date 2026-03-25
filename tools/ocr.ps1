param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]

$genericAsTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.IsGenericMethodDefinition -and
        $_.GetParameters().Count -eq 1
    } |
    Select-Object -First 1

function AwaitGeneric($Operation, [Type]$ResultType) {
    $method = $genericAsTask.MakeGenericMethod(@($ResultType))
    $task = $method.Invoke($null, @($Operation))
    return $task.GetAwaiter().GetResult()
}

$storageFileType = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$randomAccessStreamType = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$bitmapDecoderType = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$softwareBitmapType = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$ocrResultType = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]

$imageFile = AwaitGeneric ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) $storageFileType
$stream = AwaitGeneric ($imageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) $randomAccessStreamType
$decoder = AwaitGeneric ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) $bitmapDecoderType
$bitmap = AwaitGeneric ($decoder.GetSoftwareBitmapAsync()) $softwareBitmapType
$ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()

if ($null -eq $ocrEngine) {
    throw "无法初始化 Windows OCR 引擎"
}

$ocrResult = AwaitGeneric ($ocrEngine.RecognizeAsync($bitmap)) $ocrResultType
$text = if ($null -eq $ocrResult.Text) { "" } else { $ocrResult.Text }
[System.IO.File]::WriteAllText($OutputPath, $text, [System.Text.Encoding]::UTF8)
