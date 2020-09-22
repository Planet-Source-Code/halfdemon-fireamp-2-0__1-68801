FireAMP! Version 2.0.193
------------------------

[1] Features:

. Utilizes whatever codec is present in client
. Supported formats: MIDI, WAV, MP3, RM, RMVB, MOV, WMA, WMV (see notes below)
. Plugin based Visualzations
. Mass MP3 ID3 version 1.x tag editor
. View Clip properties and Video Codecs

[2] Setting up:

. Compile all dlls in the \Orange Soda\Engines\ directory and move the compiled dlls into the \Release\ directory
. Compile the FireSkinLibrary dll from the \Utils\FireScript\ Directory and move   the dll into the \Release\ directory
. Run the Setup.bat file (or register all dlls manually)
. Compile the main exe, Fireamp.vbp and move the exe to the \Release\ Directory

Setup.bat (just in case PSC removes this file)

---------------cut here-------------------
echo Registering FireAMP! Components

regsvr32 FireSkinLibrary.dll
regsvr32 Bars.dll
regsvr32 VisualizationWurx.dll
regsvr32 Particle.dll
---------------end cut--------------------

Just save the above code as a batch file and place it in the \Release\ directory.

[3] First Run:

. Visualizations require Stereo mix/Mono mix to be present on the client
. If the "Setup Stereo Mix" Option does not work, enable Stereo Mix manually
  [sndvol32.exe -> Options -> Properties -> Recording -> Select Stereo/Mono mix -> OK -> Adjust volume to just above zero]

[4] Media Support:

Real Media support requires Media Player Classic to be installed
http://filehippo.com/download_real_alternative/

Quicktime support requires Quicktime [not tested on other machines]
http://filehippo.com/download_quicktime_player/

[5] Visualizations:

Right click non-video window of the player to access the Visualization options

[6] Mass Tag Editor:

. Supports MP3 1.x tags read/write/strip
. Mass tag editor can use File Name Masks:
  %TITLE% %ARTIST% %ALBUM% %TRACK%

Eg: [06] Kuroi Namida.mp3
    Mask: [%TRACK%] %TITLE%

    01 - Jukai - Koibito Doushi
    Mask: %TRACK%- %ARTIST%- %TITLE% (notice the placement of the "-")

[7] Credits:

K.V.Rohit: for the original ID3 1.x tag reader code and lots of interface design (and testing)
Murphy McCauley (MurphyMc@Concentric.NET): for his VB FFT code (used in the Bars visualziation pack and for the wavein functions)

[8] Easter Eggs:

. Media Library
. Update over internet
. Full Screen Visualizations

These two features are not officially supported and the Update feature probably dosent even work anymore. (need to find a dedicated hosting site that does not delete your stuff for no reason)

. Keyboard shortcuts:
To put it simply, i've forgotten which shortcuts do what :) They're are a carryover from version 1.0

[9] Release Notes:

. Preferences are fully implemented yet
. Visualizations might not work the first time. Restarting the player after selecting a visualization seems to work
. Renaming files throught the Tag Editor might not reflect on the playist
. Playlist does not autoupdate
. Some video formats play outside the player [add that extension in the video playing code to make it play in the player]
. Full screen visualizations sometimes take up 100% of CPU and run slowy
. some high bitrate mp3s [320kbps typically] donot play automatically. Pressing the seek bar causes them to play.
. The player reports wrong bitrate sometimes
. Visualizations are not documented as of now.
. Skins are not documented as of now.