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