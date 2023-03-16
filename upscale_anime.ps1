$input_file_path = $args[0]
if ((Test-Path -Path $input_file_path) -ne $true) {
    Write-Host "Can't find file ${input_file_path}, skipping upscale"
    Exit
}
$input_file = Get-Item $input_file_path

#Accepts a Job as a parameter and writes the latest progress of it
function WriteJobProgress
{
    param($Job)
 
    #Make sure the first child job exists
    if($Job.ChildJobs[0].Progress -ne $null)
    {
        #Extracts the latest progress of the job and writes the progress
        $jobProgressHistory = $Job.ChildJobs[0].Progress;
        $latestProgress = $jobProgressHistory[$jobProgressHistory.Count - 1];
        $latestPercentComplete = $latestProgress | Select -expand PercentComplete;
        $latestActivity = $latestProgress | Select -expand Activity;
        $latestStatus = $latestProgress | Select -expand StatusDescription;
    
        #When adding multiple progress bars, a unique ID must be provided. Here I am providing the JobID as this
        Write-Progress -Id $Job.Id -Activity $latestActivity -Status $latestStatus -PercentComplete $latestPercentComplete;
    }
}

$begin_stopwatch = [System.Diagnostics.Stopwatch]::new()
$begin_stopwatch.Start()
$basename = $input_file.BaseName
$extension = $input_file.Extension
Write-Host "==========================="
Write-Host "Upscaling ${basename}"
Write-Host "==========================="
Write-Host "Cleaning up old files..."
if ((Test-Path ".\tmp_frames") -eq $true) {
    Remove-Item ".\tmp_frames" -Recurse > $null
}
if ((Test-Path ".\out_frames") -eq $true) {
    Remove-Item ".\out_frames" -Recurse > $null
}
Clear-RecycleBin -Force > $null
New-Item -Path ".\tmp_frames" -ItemType Directory -Force > $null
New-Item -Path ".\out_frames" -ItemType Directory -Force > $null
New-Item -Path ".\upscaled_videos" -ItemType Directory -Force > $null
$framecount = mediainfo --Output="Video;%FrameCount%" $input_file
$framerate = mediainfo --Output="Video;%FrameRate%" $input_file
Write-Host "Converting video with ${framecount} frames at ${framerate} fps"
$export_id = Get-Random
& ffmpeg -i $input_file -qscale:v 1 -qmin 1 -qmax 1 -fps_mode passthrough -v warning -stats tmp_frames/frame%08d.png 2>&1 | %{
    $found = $_ -match "frame=[ \t]*([0-9]+)[ \t]*fps=[ \t]*([0-9.]+)"
    if ($found) {
        $current_frame = [int]$matches[1]
        $current_fps = $matches[2]
        $percent = ($current_frame / $framecount) * 100
        $percent_str = ('{0:0.#}' -f $percent)
        Write-Progress -Id $export_id -Activity "(1/3) Exporting video to PNGs: ${basename}" -Status "${percent_str}% Complete (${current_fps} fps)" -PercentComplete $percent
    }
}


# Create blank files for ffmpeg
1..$framecount | % {
    $pad_num = ([string]$_).PadLeft(8,'0')
    Set-Content -Path "out_frames/frame${pad_num}.png" -value $null
}


$jobUpscalePNGs = Start-Job –Name upscale –Scriptblock {
    Param(
        $basename,
        $extension,
        $framecount,
        $framerate,
        $input_file
    )
    $upscale_stopwatch = [System.Diagnostics.Stopwatch]::new()
    $upscale_stopwatch.Start()
    $esrgan_progress = 0
    $frametime_queue = New-Object System.Collections.Queue
    $frametime_total = 0
    & realesrgan-ncnn-vulkan -i tmp_frames -o out_frames -n realesr-animevideov3 -s 2 -j 2:4:4 -f png -v 2>&1 | %{
        if ($_ -match "done$") {
            $frame_secs = $upscale_stopwatch.Elapsed.TotalSeconds
            $upscale_stopwatch.Restart()
            $frametime_total += $frame_secs
            $frametime_queue.Enqueue($frame_secs)
            while ($frametime_queue.Count -gt 16) {
                $frametime_total -= [float]($frametime_queue.Dequeue())
            }
            $fps = $frametime_queue.Count / $frametime_total
            $esrgan_progress = $esrgan_progress + 1
            $percent = ($esrgan_progress / $framecount) * 100
            $percent_str = ('{0:0.#}' -f $percent)
            $fps_str = ('{0:0.#}' -f $fps)
            $remaining_secs = ($framecount - $esrgan_progress) / $fps
            $remaining_timespan = [timespan]::fromseconds($remaining_secs)
            $remaining_str = $remaining_timespan.ToString("hh'h:'mm'm:'ss's'")
            Write-Progress -Activity "(2/3) Upscaling PNGs: ${basename}" -Status "${percent_str}% (${fps_str} fps - ${remaining_str})" -PercentComplete $percent
        }
    }
    $upscale_stopwatch.Stop()
} -ArgumentList $basename,$extension,$framecount,$framerate,$input_file


$wait_ffmpeg_stopwatch = [System.Diagnostics.Stopwatch]::new()
$wait_ffmpeg_stopwatch.Start()
while (($jobUpscalePNGs.State -ne "Completed") -and ($wait_ffmpeg_stopwatch.Elapsed.TotalMinutes -lt 3)) {
    WriteJobProgress($jobUpscalePNGs);
    Start-Sleep -Seconds 1
}
$wait_ffmpeg_stopwatch.Stop()


$jobFfmpeg = Start-Job –Name ffmpeg –Scriptblock {
    Param(
        $basename,
        $extension,
        $framecount,
        $framerate,
        $input_file
    )
    & ffmpeg -y -framerate $framerate -i out_frames/frame%08d.png -i $input_file -map 0:v:0 -map 1:a:0 -c:a copy -c:v libx265 -preset slow -crf 18 -r $framerate -pix_fmt yuv420p10le -x265-params profile=main10:bframes=8:psy-rd=1:aq-mode=3 -v warning -stats "upscaled_videos/${basename}_upscaled${extension}" 2>&1 | %{
        $found = $_ -match "frame=[ \t]*([0-9]+)[ \t]*fps=[ \t]*([0-9.]+)"
        if ($found) {
            $current_frame = [int]$matches[1]
            $fps = [float]$matches[2]
            $percent = ($current_frame / $framecount) * 100
            $percent_str = ('{0:0.#}' -f $percent)
            $fps_str = ('{0:0.#}' -f $fps)
            $remaining_secs = ($framecount - $current_frame) / $fps
            $remaining_timespan = [timespan]::fromseconds($remaining_secs)
            $remaining_str = $remaining_timespan.ToString("hh'h:'mm'm:'ss's'")
            Write-Progress -Activity "(3/3) Encoding to video file: ${basename}" -Status "${percent_str}% (${fps_str} fps - ${remaining_str})" -PercentComplete $percent
        }
    }
} -ArgumentList $basename,$extension,$framecount,$framerate,$input_file


while (($jobUpscalePNGs.State -ne "Completed") -or ($jobFfmpeg.State -ne "Completed")) {
    if ($jobUpscalePNGs.State -ne "Completed") {
        WriteJobProgress($jobUpscalePNGs);
    }
    if ($jobFfmpeg.State -ne "Completed") {
        WriteJobProgress($jobFfmpeg);
    }
 
    Start-Sleep -Seconds 1
}


$elapsed_formatted = $begin_stopwatch.Elapsed.ToString("hh'h:'mm'm:'ss's'")
$begin_stopwatch.Stop()
Write-Host "Upscale completed in ${elapsed_formatted}"
