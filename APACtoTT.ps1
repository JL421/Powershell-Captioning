<#
.SYNOPSIS
	Converts Adobe Premiere .xml files to Timed Text .xml files.
.DESCRIPTION
	Watches a specified folder for new Adobe Premiere Analyze Content XML files to convert into Timed Text 1.0 XML files and Premiere Pro 608/608 for captions.
.NOTES
	File Name	: APACtoTT.ps1
	Author		: /u/JL421
.EXAMPLE
	APACtoTT.ps1 -Watch "Path To Watch Folder" -Output "Path to Output Folder"
#>

#Get Watch and Output Folders
[cmdletbinding()]
param([string]$Watch, [string]$Output)
	
#Get all XML input files to convert	
$Inputs = Get-ChildItem -Path $Watch -filter "*.xml"

#Main Function

foreach ($Input in $Inputs)
{
IF ($Input -ne $Null)
{
	$SMPTETTFile = '<?xml version="1.0" encoding="utf-8"?>
<tt xmlns="http://www.w3.org/ns/ttml" xmlns:ttp="http://www.w3.org/ns/ttml#parameter" ttp:timeBase="media" xmlns:tts="http://www.w3.org/ns/ttml#style" xml:lang="en" xmlns:ttm="http://www.w3.org/ns/ttml#metadata">
  <head>
    <metadata>
      <ttm:title></ttm:title>
    </metadata>
    <styling>
      <style id="s0" tts:backgroundColor="black" tts:fontStyle="normal" tts:fontSize="16" tts:fontFamily="sansSerif" tts:color="white" />
    </styling>
  </head>
  <body style="s0">
    <div>'
	
	$608File = '<?xml version="1.0" encoding="UTF-8" standalone="no" ?>
<tt xmlns="http://www.w3.org/ns/ttml" xmlns:ttp="http://www.w3.org/ns/ttml#parameter" ttp:dropMode="dropNTSC" ttp:frameRate="30" ttp:frameRateMultiplier="1000 1001" ttp:timeBase="smpte" xmlns:m608="http://www.smpte-ra.org/schemas/2052-1/2013/smpte-tt#cea608" xmlns:smpte="http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt" xmlns:ttm="http://www.w3.org/ns/ttml#metadata" xmlns:tts="http://www.w3.org/ns/ttml#styling">

  <head>
    <styling>
      <style tts:color="white" tts:fontFamily="monospace" tts:fontWeight="normal" tts:textAlign="left" xml:id="basic"/>
    </styling>
    <layout>
      <region tts:backgroundColor="transparent" xml:id="pop1"/>
    </layout>
    <metadata/>
    <smpte:information m608:aspectRatio="16:9" m608:easyReader="false" m608:number="1"/>
  </head>
  <body>
    <div>'
	
		#Import Adobe Analyze Content file into memory
		
		if ($Input.Attributes -ne "Directory")
		{
			$InputXML = New-Object 'object[,]' 0,0
			$SMPTETTXML = New-Object 'object[,]' 0,0
			$608XML = New-Object 'object[,]' 0,0

			$Time = $null
			$Content = $null
			$Duration = $null
			$Confidence = $null
			[xml]$InputContent = Get-Content $Watch\$Input
			foreach ($CuePoint in $InputContent.FLVCoreCuePoints.CuePoint)
			{
				$Time = $CuePoint.Time
				$Content = $CuePoint.Name
				foreach ($Parameter in $CuePoint.Parameters.Parameter)
				{
					if ($Parameter.Name -eq "duration")
					{
						$Duration = $Parameter.Value
					}
					if ($Parameter.Name -eq "confidence")
					{
						[int]$Confidence = $Parameter.Value
					}
				}
				if ($Time -ne $Null)
				{
					IF ($Confidence -eq 0)
					{
						Continue
					}
					$InputXML += ,@($Time,$Content,$Duration,$Confidence)
				}
			}
			
			#SMPTETT Process
			
			$Counter = 0
			
			do {
				$WorkingSetCurrent = $InputXML[$Counter]
	
				$CurrentTime = [int]$WorkingSetCurrent[0]
				$CurrentString = [string]$WorkingSetCurrent[1]
				$CurrentDuration = [int]$WorkingSetCurrent[2]
				$CurrentConfidence = [int]$WorkingSetCurrent[3]
				$CurrentNumber = [int]$Counter
				
				$MaxString = $FALSE
				
				$CountNext = $Counter
				do {
					$CountNext++
					if ($CountNext -eq $InputXML.Count)
					{
						$Counter = $InputXML.Count
						break
					}
					$WorkingSetNext = $InputXML[$CountNext]
					$NextTime = [int]$WorkingSetNext[0]
					$NextString = [string]$WorkingSetNext[1]
					$NextDuration = [int]$WorkingSetNext[2]
					$NextConfidence = [int]$WorkingSetNext[3]
					$StringLength = $CurrentString + " " + $NextString
					$StringLength = $StringLength | Measure-Object -Character | select -expandproperty characters
	
					if ($CurrentTime + $CurrentDuration + 500 -ge $NextTime -And $StringLength -le 43)
					{
						$CurrentString = $CurrentString + " " + $NextString
						$CurrentDuration = $CurrentDuration + $NextDuration
						$CurrentConfidence = $CurrentConfidence + $NextConfidence
					} elseif($CurrentTime + $CurrentDuration + 500 -ge $NextTime -And $StringLength -gt 43)
					{
						$CurrentDuration = $NextTime - $CurrentTime
						$MaxString = $TRUE
					} elseif ($CurrentTime + $CurrentDuration + 500 -le $NextTime -And $CurrentDuration -le 1000)
					{
						if ($CurrentTime + $CurrentDuration + 1000 -ge $NextTime)
						{
							$CurrentDuration = $NextTime - $CurrentTime
						} else {
							$CurrentDuration = $CurrentDuration + 1000
						}
						$MaxString = $TRUE
					} elseif ($CurrentTime + $CurrentDuration + 500 -le $NextTime -And $CurrentDuration -gt 1000)
					{
						if ($CurrentTime + $CurrentDuration + 1000 -ge $NextTime)
						{
							$CurrentDuration = $NextTime - $CurrentTime
						} else {
							$CurrentDuration = $CurrentDuration + 1000
						}
						$MaxString = $TRUE
					} else {
						$MaxString = $TRUE
					}
					$Counter = $CountNext
				}
				while ($MaxString -eq $FALSE -And $CountNext -le $InputXML.Count)
				if ($CurrentString -ne "AND" -And $CurrentString -ne "WOAH" -And $CurrentString -ne "WHOA" -And $CurrentString -ne "IT WAS")
				{
					#$NumberofWords = (($CountNext - 1) - $CurrentNumber) + 1
					#$AverageConfidence = $CurrentConfidence/$NumberofWords
					IF ($CurrentDuration/$NumberofWords -lt 1000)
					{
						$CurrentString = $CurrentString.ToUpper()
						$SMPTETTXML += ,@($CurrentTime,$CurrentString,$CurrentDuration,$AverageConfidence)
					}
				}
			}
			Until ($Counter -eq $InputXML.Count)
	
			#608 Process
			
			$Counter = 0
			
			do {
				$WorkingSetCurrent = $InputXML[$Counter]
	
				$CurrentTime = [int]$WorkingSetCurrent[0]
				$CurrentString = [string]$WorkingSetCurrent[1]
				$CurrentDuration = [int]$WorkingSetCurrent[2]
				
				$MaxString = $FALSE
				
				$CountNext = $Counter
				do {
					$CountNext++
					if ($CountNext -eq $InputXML.Count)
					{
						$Counter = $InputXML.Count
						break
					}
					$WorkingSetNext = $InputXML[$CountNext]
					$NextTime = [int]$WorkingSetNext[0]
					$NextString = [string]$WorkingSetNext[1]
					$NextDuration = [int]$WorkingSetNext[2]
					$StringLength = $CurrentString + " " + $NextString
					$StringLength = $StringLength | Measure-Object -Character | select -expandproperty characters
	
					if ($CurrentTime + $CurrentDuration + 500 -ge $NextTime -And $StringLength -le 32)
					{
						$CurrentString = $CurrentString + " " + $NextString
						$CurrentDuration = $CurrentDuration + $NextDuration
					} elseif($CurrentTime + $CurrentDuration + 500 -ge $NextTime -And $StringLength -gt 32)
					{
						$CurrentDuration = $NextTime - $CurrentTime
						$MaxString = $TRUE
					} elseif ($CurrentTime + $CurrentDuration + 500 -le $NextTime -And $CurrentDuration -le 1000)
					{
						if ($CurrentTime + $CurrentDuration + 1000 -ge $NextTime)
						{
							$CurrentDuration = $NextTime - $CurrentTime
						} else {
							$CurrentDuration = $CurrentDuration + 1000
						}
						$MaxString = $TRUE
					} elseif ($CurrentTime + $CurrentDuration + 500 -le $NextTime -And $CurrentDuration -gt 1000)
					{
						if ($CurrentTime + $CurrentDuration + 1000 -ge $NextTime)
						{
							$CurrentDuration = $NextTime - $CurrentTime
						} else {
							$CurrentDuration = $CurrentDuration + 1000
						}
						$MaxString = $TRUE
					} else {
						$MaxString = $TRUE
					}
					$Counter = $CountNext
				}
				while ($MaxString -eq $FALSE -And $CountNext -le $InputXML.Count)
				if ($CurrentString -ne "AND" -And $CurrentString -ne "WOAH" -And $CurrentString -ne "WHOA" -And $CurrentString -ne "IT WAS")
				{
					$CurrentString = $CurrentString.ToUpper()
					$608XML += ,@($CurrentTime,$CurrentString,$CurrentDuration)
				}
			}
			Until ($Counter -eq $InputXML.Count)
			
			#Build SMPTETT Lines
			
			$Counter = 0
			
			foreach ($i in $SMPTETTXML)
			{
				$Caption = $i[1]
				$StartTimeMS = $i[0]
				$EndTimeMS = $i[0] + $i[2]
				$StartTimeFrame = [System.Math]::Floor([System.Math]::Round($StartTimeMS / 1000 - [System.Math]::Floor($StartTimeMS / 1000), 4) * 1000 / (1000 / 24))
				$StartSeconds = [System.Math]::Floor($StartTimeMS / 1000)
				$StartTimeSpan = New-TimeSpan -Seconds $StartSeconds
				$StartTime = "{0:D2}:{1:D2}:{2:D2}" -f $StartTimeSpan.Hours,$StartTimeSpan.Minutes,$StartTimeSpan.Seconds + ":" + "{0:D2}" -f [int]$StartTimeFrame
				$EndTimeFrame = [System.Math]::Floor([System.Math]::Round($EndTimeMS / 1000 - [System.Math]::Floor($EndTimeMS / 1000), 4) * 1000 / (1000 / 24))
				$EndSeconds = [System.Math]::Floor($EndTimeMS / 1000)
				$EndTimeSpan = New-TimeSpan -Seconds $EndSeconds
				$EndTime = "{0:D2}:{1:D2}:{2:D2}" -f $EndTimeSpan.Hours,$EndTimeSpan.Minutes,$EndTimeSpan.Seconds + ":" + "{0:D2}" -f [int]$EndTimeFrame
				$SMPTETTString = '      <p begin="' + $StartTime + '" id="p' + $Counter + '" end="' + $EndTime + '">' + $Caption + '</p>' # + '  <!-- Confidence is ' + $i[3] + '  -->'
				$SMPTETTFile = $SMPTETTFile + "`n" + $SMPTETTString
				$Counter++
			}
			
			#Build 608 Lines
			
			foreach ($i in $608XML)
			{
				$Caption = $i[1]
				$StartTimeMS = $i[0]
				$EndTimeMS = $i[0] + $i[2]
				$CaptionLength = $Caption | Measure-Object -Character | select -expandproperty characters
				$HorizontalPos = 50 - ($CaptionLength * 1.25)
				$StartTimeFrame = [System.Math]::Floor([System.Math]::Round($StartTimeMS / 1000 - [System.Math]::Floor($StartTimeMS / 1000), 4) * 1000 / (1000 / 24))
				$StartSeconds = [System.Math]::Floor($StartTimeMS / 1000)
				$StartTimeSpan = New-TimeSpan -Seconds $StartSeconds
				$StartTime = "{0:D2}:{1:D2}:{2:D2}" -f $StartTimeSpan.Hours,$StartTimeSpan.Minutes,$StartTimeSpan.Seconds + ":" + "{0:D2}" -f [int]$StartTimeFrame
				$EndTimeFrame = [System.Math]::Floor([System.Math]::Round($EndTimeMS / 1000 - [System.Math]::Floor($EndTimeMS / 1000), 4) * 1000 / (1000 / 24))
				$EndSeconds = [System.Math]::Floor($EndTimeMS / 1000)
				$EndTimeSpan = New-TimeSpan -Seconds $EndSeconds
				$EndTime = "{0:D2}:{1:D2}:{2:D2}" -f $EndTimeSpan.Hours,$EndTimeSpan.Minutes,$EndTimeSpan.Seconds + ":" + "{0:D2}" -f [int]$EndTimeFrame
				$608String1 = '      <p begin="' + $StartTime + '" end="' + $EndTime + '" region="pop1" style="basic" tts:origin="' + $HorizontalPos + '% 85%">'
				$608String2 = '        <style/>'
				$608String3 = '        <span>' + $Caption
				$608String4 = '          <style tts:backgroundColor="#000000FF" tts:color="#AAAAAAFF" tts:fontSize="18px"/>'
				$608String5 = '        </span>'
				$608String6 = '      </p>'
				$608File = $608File + "`n" + $608String1 + "`n" + $608String2 + "`n" + $608String3 + "`n" + $608String4 + "`n" + $608String5 + "`n" + $608String6
			}
		}
	
	#Write out SMPTETT File
	
	$SMPTETTFile = $SMPTETTFile + '
    </div>
  </body>
</tt>'
		$SMPTETTFile = $SMPTETTFile -replace "`n", "`r`n"
		$SMPTETTPath = "$Output\SMPTETT.xml"
		[System.IO.File]::WriteAllLines($SMPTETTPath, $SMPTETTFile)
		Remove-Item $Watch\$Input
	#Write out 608 File
	
	$608File = $608File + '
    </div>
  </body>
</tt>'
		$608File = $608File -replace "`n", "`r`n"
		$608Path = "$Output\608.xml"
		[System.IO.File]::WriteAllLines($608Path, $608File)
		
	}
}