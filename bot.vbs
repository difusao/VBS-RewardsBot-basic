Set WshShell = WScript.CreateObject("WScript.Shell")
Set objRandom = CreateObject( "System.Random" )

Randomize

Dim sEdge: sEdge = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
Dim sArrProfile : sArrProfile = Array("""--profile-directory=default""")

WshShell.Run sEdge & " " & sArrProfile(0) & " https://rewards.bing.com/", 0, TRUE
WScript.Sleep(5000)
	
For i = 0 To 30
	sUrl = "https://www.bing.com/search?q=""o esporte nos anos de 19" & Int((99 * Rnd) + 1)
	WScript.Sleep(3000)
	WshShell.Run sEdge & " " & sArrProfile(0) & " " & sUrl, 0, TRUE
Next

WScript.Sleep(iIntervalSearch)
WshShell.Run sEdge & " " & sArrProfile(0) & " https://rewards.bing.com/status/pointsbreakdown", 0, TRUE