 robocopy 

 if ($lastexitcode -eq 0)
 {
      write-host "Robocopy succeeded"
 }

 if ($lastexitcode -eq 1)
 {
      write-host "Robocopy succeeded with exit code 1: All files were copied successfully" 
 }

 if ($lastexitcode -eq 2) 
 {
      Write-Host "Robocopy failed with exit code 2: There are some additional files in the destination directory that are not present in the source directory. No files were copied."
 }

 if ($lastexitcode -eq 3)
 {
      Write-Host "Some files were copied. Additional files were present. No failure was encountered"
 }

 if ($lastexitcode -eq 5)
 {
      Write-Host "Some files were copied. Some files were mismatched. No failure was encountered."
 }

 if ($lastexitcode -eq 6)
 {
      Write-Host "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory."
 }

 if ($lastexitcode -eq 7)
 {
      Write-Host "Files were copied, a file mismatch was present, and additional files were present."
 }

 if ($lastexitcode -eq 8)
 {
      Write-Host "Robocopy failed with exit code 8: Several files did not copy."
 }

else
{
      write-host "Robocopy failed with exit code:" $lastexitcode
}