param (
    [Parameter(Mandatory=$true)][string]$input_file_path_unescaped,
    [Int32]$output_height = -2,
    [Int32]$scale = 2
)

$input_file_path = [Management.Automation.WildcardPattern]::Escape($input_file_path_unescaped)
if ((Test-Path -Path $input_file_path) -ne $true) {
    Write-Host "Can't find file ${input_file_path}, skipping upscale"
    Exit
}
$input_file = Get-Item $input_file_path

if (($scale -ne 2) -and ($scale -ne 3) -and ($scale -ne 4)) {
    Write-Host "Scale of ${scale} not supported (must be 2,3,4), skipping upscale"
    Exit
} else {
    Write-Host "Using scale of ${scale}"
}

if ($output_height % 2 -ne 0) {
    Write-Host "Output height of ${output_height} not supported (must be even), skipping upscale"
    Exit
} else {
    Write-Host "Using output height of ${output_height}"
}


#Accepts a Job as a parameter and writes the latest progress of it
function WriteJobProgress($Job, $JobId) {
    #Make sure the first child job exists
    if($Job.ChildJobs[0].Progress -ne $null)
    {
        #Extracts the latest progress of the job and writes the progress
        $jobProgressHistory = $Job.ChildJobs[0].Progress;
        $latestProgress = $jobProgressHistory[$jobProgressHistory.Count - 1];
        $latestPercentComplete = $latestProgress | Select -expand PercentComplete;
        $latestActivity = $latestProgress | Select -expand Activity;
        $latestStatus = $latestProgress | Select -expand StatusDescription;

        if ($JobId -eq $null) {
            $JobId = $Job.Id
        }
    
        #When adding multiple progress bars, a unique ID must be provided. Here I am providing the JobID as this
        Write-Progress -Id $JobId -Activity $latestActivity -Status $latestStatus -PercentComplete $latestPercentComplete;
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
New-Item -Path ".\tmp_frames" -ItemType Directory -Force > $null
New-Item -Path ".\out_frames" -ItemType Directory -Force > $null
New-Item -Path ".\upscaled_videos" -ItemType Directory -Force > $null
$framecount = ffprobe -count_frames -v error -select_streams v:0 -show_entries stream=nb_read_frames -of default=nokey=1:noprint_wrappers=1 -threads 24 $input_file
$framerate = ffprobe -v error -select_streams v:0 -show_entries stream=r_frame_rate -of default=nokey=1:noprint_wrappers=1 $input_file
Write-Host "Converting video with ${framecount} frames at ${framerate} fps"


# Create blank files for upscaling/ffmpeg
$generate_blanks_id = Get-Random
$generate_blanks_stopwatch = [System.Diagnostics.Stopwatch]::new()
$generate_blanks_stopwatch.Start()
for ($i=1; $i -le $framecount; $i++) {
    $pad_num = ([string]$i).PadLeft(8,'0')
    Set-Content -Path "tmp_frames/frame${pad_num}.png" -value $null
    Set-Content -Path "out_frames/frame${pad_num}.png" -value $null
    if ($generate_blanks_stopwatch.ElapsedMilliseconds -ge 1000) {
        $percent = $i / $framecount * 100
        $percent_str = $percent_str = ('{0:0.##}' -f $percent)
        Write-Progress -Id $generate_blanks_id -Activity "(0/3) Generating blank PNGs: ${basename}" -Status "${percent_str}% Complete" -PercentComplete $percent
        $generate_blanks_stopwatch.Restart()
    }
}
$generate_blanks_stopwatch.Stop()
Write-Progress -Id $generate_blanks_id -Activity "(0/3) Generating blank PNGs: ${basename}" -Status "100% Complete" -PercentComplete 100


$chunk_size = 4000


$scriptBlockExportPNGs = {
    Param(
        $basename,
        $extension,
        $framecount,
        $framerate,
        $input_file,
        $chunk_size,
        $start_index
    )
    $ffmpeg_start_index = $start_index - 1
    $ffmpeg_end_index = (($ffmpeg_start_index + $chunk_size),$framecount | Measure -Min).Minimum
    & ffmpeg -y -i $input_file -vf "trim=start_frame=${ffmpeg_start_index}:end_frame=${ffmpeg_end_index}" -start_number $start_index -qscale:v 1 -qmin 1 -qmax 1 -fps_mode passthrough -v warning -stats tmp_frames/frame%08d.png 2>&1 | %{
        $found = $_ -match "frame=[ \t]*([0-9]+)[ \t]*fps=[ \t]*([0-9.]+)"
        if ($found) {
            $current_frame = [int]$matches[1] + $ffmpeg_start_index
            $current_fps = $matches[2]
            $percent = ($current_frame / $framecount) * 100
            $percent_str = ('{0:0.##}' -f $percent)
            Write-Progress -Activity "(1/3) Exporting video to PNGs: ${basename}" -Status "${percent_str}% Complete (${current_fps} fps)" -PercentComplete $percent
        }
    }
    if ($ffmpeg_end_index -eq $framecount) {
        Write-Progress -Activity "(1/3) Exporting video to PNGs: ${basename}" -Status "100% Complete" -PercentComplete 100
    }
}


$jobExportPNGs = Start-Job –Name export1 –Scriptblock $scriptBlockExportPNGs -ArgumentList $basename,$extension,$framecount,$framerate,$input_file,$chunk_size,1
$export_pngs_completing = $chunk_size


$jobUpscalePNGs = Start-Job –Name upscale –Scriptblock {
    Param(
        $basename,
        $extension,
        $framecount,
        $framerate,
        $scale,
        $input_file
    )
    # wait 30 seconds for export to get started
    Start-Sleep -Seconds 30
    $cleanedup_until = 0
    $upscale_stopwatch = [System.Diagnostics.Stopwatch]::new()
    $upscale_stopwatch.Start()
    $esrgan_progress = 0
    $frametime_queue = New-Object System.Collections.Queue
    $frametime_total = 0
    & realesrgan-ncnn-vulkan -i tmp_frames -o out_frames -n realesr-animevideov3 -s $scale -t 4096 -j 2:4:4 -f png -v 2>&1 | %{
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
            $percent_str = ('{0:0.##}' -f $percent)
            $fps_str = ('{0:0.#}' -f $fps)
            $remaining_secs = ($framecount - $esrgan_progress) / $fps
            $remaining_timespan = [timespan]::fromseconds($remaining_secs)
            $remaining_str = $remaining_timespan.ToString("hh'h:'mm'm:'ss's'")
            if (($esrgan_progress - $cleanedup_until) -gt 500) {
                $next_cleanup_until = $esrgan_progress - 500
                ($cleanedup_until+1)..$next_cleanup_until | ForEach-Object {
                    $pad_num = ([string]$_).PadLeft(8,'0')
                    Remove-Item "tmp_frames/frame${pad_num}.png"
                }
                $cleanedup_until = $next_cleanup_until
            }
            Write-Progress -Activity "(2/3) Upscaling PNGs: ${basename}" -Status "${percent_str}% (${fps_str} fps - ${remaining_str})" -PercentComplete $percent
        }
    }
    $upscale_stopwatch.Stop()
    Write-Progress -Activity "(2/3) Upscaling PNGs: ${basename}" -Status "100%" -PercentComplete 100
} -ArgumentList $basename,$extension,$framecount,$framerate,$scale,$input_file


$jobFfmpeg = Start-Job –Name ffmpeg –Scriptblock {
    Param(
        $basename,
        $extension,
        $framecount,
        $framerate,
        $output_height,
        $input_file
    )
    # wait 3 minutes for upscale to get started
    Start-Sleep -Seconds 180
    $cleanedup_until = 0
    & ffmpeg -y -framerate $framerate -i out_frames/frame%08d.png -i $input_file -map 0:v:0 -map 1:a -map 1:s? -map_metadata 1 -map_chapters 1 -c:a copy -c:s copy -c:v libx265 -preset slow -crf 18 -r $framerate -pix_fmt yuv420p10le -x265-params profile=main10:bframes=8:psy-rd=1:aq-mode=3 -vf "scale=-2:${output_height}:filter=spline36" -v warning -stats "upscaled_videos/${basename}_upscaled.mkv" 2>&1 | %{
        $found = $_ -match "frame=[ \t]*([0-9]+)[ \t]*fps=[ \t]*([0-9.]+)"
        if ($found) {
            $current_frame = [int]$matches[1]
            $fps = [float]$matches[2]
            $percent = ($current_frame / $framecount) * 100
            $percent_str = ('{0:0.##}' -f $percent)
            $fps_str = ('{0:0.#}' -f $fps)
            $remaining_secs = ($framecount - $current_frame) / $fps
            $remaining_timespan = [timespan]::fromseconds($remaining_secs)
            $remaining_str = $remaining_timespan.ToString("hh'h:'mm'm:'ss's'")
            if (($current_frame - $cleanedup_until) -gt 500) {
                $next_cleanup_until = $current_frame - 500
                ($cleanedup_until+1)..$next_cleanup_until | ForEach-Object {
                    $pad_num = ([string]$_).PadLeft(8,'0')
                    Remove-Item "out_frames/frame${pad_num}.png"
                }
                $cleanedup_until = $next_cleanup_until
            }
            Write-Progress -Activity "(3/3) Encoding to video file: ${basename}" -Status "${percent_str}% (${fps_str} fps - ${remaining_str})" -PercentComplete $percent
        }
    }
    Write-Progress -Activity "(3/3) Encoding to video file: ${basename}" -Status "100%" -PercentComplete 100
} -ArgumentList $basename,$extension,$framecount,$framerate,$output_height,$input_file


$exportPNGsProgressId = Get-Random

$wsh = New-Object -ComObject WScript.Shell

while (($jobExportPNGs.State -ne "Completed") -or ($jobUpscalePNGs.State -ne "Completed") -or ($jobFfmpeg.State -ne "Completed")) {
    if ($jobExportPNGs.State -ne "Completed") {
        WriteJobProgress $jobExportPNGs $exportPNGsProgressId
    } else {
        if (($jobUpscalePNGs.ChildJobs[0].Progress -ne $null) -and ($export_pngs_completing -lt $framecount)) {
            $jobProgressHistory = $jobUpscalePNGs.ChildJobs[0].Progress;
            $latestProgress = $jobProgressHistory[$jobProgressHistory.Count - 1];
            $latestStatus = $latestProgress | Select -expand StatusDescription;
            $found = $latestStatus -match "^([0-9.]+)\%"
            if ($found) {
                $percent = [float]$matches[1]
                $upscale_frame = ($percent / 100) * $framecount
                if ((($percent / 100) * $framecount) -gt ($export_pngs_completing - ($chunk_size / 2))) {
                    # we need to run another batch
                    $start_index = ($export_pngs_completing + 1)
                    $jobExportPNGs = Start-Job –Name "export${start_index}" –Scriptblock $scriptBlockExportPNGs -ArgumentList $basename,$extension,$framecount,$framerate,$input_file,$chunk_size,$start_index
                    $export_pngs_completing += $chunk_size
                }
            }
        }
    }
    if ($jobUpscalePNGs.State -ne "Completed") {
        WriteJobProgress $jobUpscalePNGs
    }
    if ($jobFfmpeg.State -ne "Completed") {
        WriteJobProgress $jobFfmpeg
    }
 
    # Keep computer from sleeping
    $wsh.SendKeys('+{F15}')

    Start-Sleep -Seconds 1
}


$elapsed_formatted = $begin_stopwatch.Elapsed.ToString("hh'h:'mm'm:'ss's'")
$begin_stopwatch.Stop()
Write-Host "Upscale completed in ${elapsed_formatted}"
