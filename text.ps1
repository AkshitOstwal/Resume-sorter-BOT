
    Param(
         [parameter(position=0)]
        $path,

 [parameter(position=1)]
 $RDSL
         )
 # path is the path of folder containing Resumes
 $files    = Get-Childitem $path -Include *.docx,*.doc,*.pdf -Recurse | Where-Object { !($_.psiscontainer) }
 $output   = "c:\Users\akshi\Desktop\wordfiletry.txt"
 $application = New-Object -comobject word.application
 $application.visible = $False
 $findtext = @($RDSL.split(','))
 $Emailpattern = "(?<email>([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5}))"
 $email = ''
 $number = ''
 $PhoneNoPattern = "(?<number>([+]|)(\d{1,3}[-\s]?|)\d{3}[-\s]?\d{3}[-\s]?\d{4})"
 $directoyPath="C:\Users\akshi\Desktop\output Resumes";

      '' | Out-File  $output
      if(!(Test-Path -path $directoyPath))  
        {  
            New-Item -ItemType directory -Path $directoyPath               
        }
      else
        {
            Remove-Item "$directoyPath\*.*" | Where { ! $_.PSIsContainer }
        }
      'DO NOT CLOSE THIS WINDOW, IT WILL CLOSE AUTOMATICALLY ONCE WORK IS COMPLETED !!!'
      # Loop through all *.doc *.pdf files in the $path directory
      Foreach ($file In $files)
      {
            $document = $application.documents.open($file.FullName,$false,$true)
            $paras = $document.Paragraphs
       
            Foreach ($para in $paras)
                {
                    $EmailFound = $para.Range.Text -match $Emailpattern
                    if($EmailFound)
                    {   
                        $email = $Matches.email
                         break
                    }
                }
            Foreach ($para in $paras)
                {
                    $PhoneNoFound  = $para.Range.Text -match $PhoneNoPattern
                       
                    if($PhoneNoFound)
                    {
                       $number = $Matches.number
                        break
                    }
                 }
            Foreach ($word in $findText)
                { 
                    $range = $document.content
            
                    $wordFound = $range.find.execute($word)
            
                    if($wordFound)
                        {  
                            $file.Name + " : " +$file.BaseName + " : $word" + " : " + $email +" : " + $number | Out-File $output -Append
                            Copy-Item $file -Destination $directoyPath
                        } 
                }
            $email = ''
            $number = ''
            $document.close()
      }

    $application.quit()
    
    # powershell -ExecutionPolicy Bypass -File text.ps1 "C:\Users\akshi\Desktop\Sample Resume ( Input)" "java"
