namespace SpreadsheetApp

type App() as app =
  inherit System.Windows.Application()
  do app.Startup.Add(fun _ -> app.RootVisual <- Spreadsheet.View())