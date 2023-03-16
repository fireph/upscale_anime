# upscale_anime
Scripts to help upscale anime (using [Real-ESRGAN](https://github.com/xinntao/Real-ESRGAN))

Can run on many files:
```
Get-ChildItem "." -Filter *.mkv | 
Foreach-Object {
    upscale_anime $_
}
```
