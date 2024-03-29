'Install .Net 3.5 via Control Panel, Program and Features, Turn Windows features on or off

'Download and install Selenium Basic from: https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0

'Download the correct Chrome web driver from: https://chromedriver.chromium.org/downloads

'Move the new Chrome web driver to %LocalAppData%\SeleniumBasic, replacing the old version

'Save this script anywhere as something like SeleniumBasicTest.vbs and just double-click it

Set SeleniumDriver = CreateObject("Selenium.ChromeDriver")
With SeleniumDriver
  .Get "https://cbracco.github.io/html5-test-page/"
  .FindElementByID("input__text").SendKeys("Pikachu")
  WScript.Sleep 1000
  .FindElementByID("input__emailaddress").SendKeys("Pikachu@pokemon.com")
  WScript.Sleep 1000
  .FindElementByID("select").SendKeys(.Keys.ENTER & .Keys.DOWN & .Keys.DOWN)
  WScript.Sleep 1000
  .FindElementByID("checkbox2").SendKeys(.Keys.Space)
  .FindElementByID("checkbox3").SendKeys(.Keys.Space)
  WScript.Sleep 1000
  .FindElementByID("radio3").SendKeys(.Keys.Space)
  WScript.Sleep 1000
  For i = 1 To 5
    .FindElementByID("textarea").SendKeys("The quick brown fox jumps over the lazy dog" & VBCRLF)
  Next
 '.Quit
End With
WScript.Sleep 10000