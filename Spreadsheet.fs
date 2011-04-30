﻿#if INTERACTIVE
#else
module Spreadsheet
#endif

open System
open System.Collections
open System.ComponentModel
open System.Windows
open System.Windows.Controls
open System.Windows.Data
open System.Windows.Input
open System.Windows.Media

let rec toRef (s:string) =
    let col = int s.[0] - int 'A'
    let row = s.Substring 1 |> int
    col, row-1

type token =
    | WhiteSpace
    | Symbol of char
    | OpToken of string
    | RefToken of int * int
    | FunToken of string
    | NumToken of decimal 

let (|Match|_|) pattern input =
    let m = System.Text.RegularExpressions.Regex.Match(input, pattern)
    if m.Success then Some m.Value else None

let matchToken = function
    | Match @"^\s+" s -> s, WhiteSpace
    | Match @"^\+"  s -> s, OpToken s
    | Match @"^\-"  s -> s, OpToken s
    | Match @"^\*"  s -> s, OpToken s
    | Match @"^\/"  s -> s, OpToken s
    | Match @"^\(|^\)|^\," s -> s, Symbol s.[0]   
    | Match @"^[A-Z]\d+" s -> s, s |> toRef |> RefToken
    | Match @"^[A-Za-z]+" s -> s, FunToken s
    | Match @"^\d+(\.\d+)?|\.\d+" s -> s, s |> decimal |> NumToken
    | _ -> invalidOp ""

let tokenize s =
    let rec tokenize' index (s:string) =
        if index = s.Length then [] 
        else
            let next = s.Substring index 
            let text, token = matchToken next
            token :: tokenize' (index + text.Length) s
    tokenize' 0 s
    |> List.choose (function WhiteSpace -> None | t -> Some t)

type operator = Add | Sub | Mul | Div

type expr =
    | Neg of expr
    | Op of expr * operator * expr
    | Num of decimal
    | Ref of int * int
    | Fun of string * expr list

let rec (|Term|_|) = function  
    | Factor(e1, t) ->      
        let rec aux e1 = function        
            | Sum op::Factor(e2, t) -> aux (Op(e1,op,e2)) t               
            | t -> Some(e1, t)      
        aux e1 t  
    | _ -> None
and (|Sum|_|) = function 
    | OpToken "+" -> Some Add 
    | OpToken "-" -> Some Sub 
    | _ -> None
and (|Factor|_|) = function  
    | OpToken "-"::Factor(e, t) -> Some(Neg(e), t)  
    | Atom(e1, Product op::Factor(e2, t)) ->
        Some(Op(e1,op,e2), t)       
    | Atom(e, t) -> Some(e, t)  
    | _ -> None    
and (|Product|_|) = function
    | OpToken "*" -> Some Mul
    | OpToken "/" -> Some Div
    | _ -> None
and (|Atom|_|) = function      
    | NumToken(n)::t -> Some(Num(n), t)   
    | RefToken(x,y)::t -> Some(Ref(x,y), t)
    | Symbol '('::Term(e, Symbol ')'::t) -> Some(e, t)
    | FunToken(s)::Tuple(ps, t) -> Some(Fun(s,ps),t)  
    | _ -> None
and (|Tuple|_|) = function
    | Symbol '('::Params(ps, Symbol ')'::t) -> Some(ps, t)  
    | _ -> None
and (|Params|_|) = function
    | Term(e1, t) ->
        let rec aux es = function
            | Symbol ','::Term(e2, t) -> aux (es@[e2]) t
            | t -> es, t
        Some(aux [e1] t)
    | t -> Some ([],t)

let parse s = 
    tokenize s |> function 
    | Term(e,[]) -> e 
    | _ -> failwith "Failed to parse formula"

let evaluate valueAt s =
    let rec eval = function
        | Neg e -> - (eval e)
        | Op(e1,Add,e2) -> eval e1 + eval e2
        | Op(e1,Sub,e2) -> eval e1 - eval e2
        | Op(e1,Mul,e2) -> eval e1 * eval e2
        | Op(e1,Div,e2) -> eval e1 / eval e2
        | Num d -> d
        | Ref(x,y) -> valueAt(x,y) |> decimal
        | Fun("SUM",ps) -> ps |> List.map eval |> List.sum
        | Fun(_,_) -> failwith "Unknown function"
    eval s

let references s =
    let rec traverse = function
        | Ref(x,y) -> [x,y]
        | Fun(_,ps) -> ps |> List.collect traverse
        | Op(e1,_,e2) -> traverse e1 @ traverse e2
        | _ -> []
    traverse s

type Sheet (colCount,rowCount) as sheet =
    let cols = Array.init colCount (fun i -> string (int 'A' + i |> char)) 
    let rows = Array.init rowCount (fun index -> Row(index+1,colCount,sheet))  
    member sheet.Columns = cols    
    member sheet.Rows = rows
    member sheet.MaxGeneration = 1000
and Row (index,colCount,sheet) =
    let cells = Array.init colCount (fun i -> Cell(sheet))
    member row.Cells = cells
    member row.Index = index
and Cell (sheet:Sheet) as cell =
    inherit ObservableObject()
    let mutable value = ""
    let mutable data = ""       
    let mutable expr : expr option = None
    let updated = Event<_>()
    let mutable subscriptions : System.IDisposable list = []   
    let cellAt(x,y) = 
        let (row : Row) = Array.get sheet.Rows y
        let (cell : Cell) = Array.get row.Cells x
        cell
    let valueAt address = (cellAt address).Value
    let eval expr =         
        try (evaluate valueAt expr).ToString()       
        with _ -> "N/A"
    let parseFormula (text:string) =
        if text.StartsWith "="
        then                
            try true, parse (text.Substring 1) |> Some
            with _ -> true, None
        else false, None
    let update newValue generation =
        if newValue <> value then
            value <- newValue
            updated.Trigger generation
            cell.Notify "Value"
    let unsubscribe () =
        subscriptions |> List.iter (fun d -> d.Dispose())
        subscriptions <- []
    let subscribe expr addresses =
        let remember x = subscriptions <- x :: subscriptions
        for address in addresses do
            let cell' : Cell = cellAt address
            cell'.Updated
            |> Observable.subscribe (fun generation ->   
                if generation < sheet.MaxGeneration then
                    let newValue = eval expr
                    update newValue (generation+1)
            ) |> remember
    member cell.Data 
        with get () = data 
        and set (text:string) =
            data <- text        
            cell.Notify "Data"          
            let isFormula, newExpr = parseFormula text               
            expr <- newExpr
            unsubscribe()
            expr |> Option.iter (fun expr ->  references expr |> subscribe expr)
            let newValue =
                match isFormula, expr with           
                | _, Some expr -> eval expr
                | true, _ -> "N/A"
                | _, None -> text
            update newValue 0
    member cell.Value = value
    member cell.Updated = updated.Publish
and ObservableObject() =
    let propertyChanged = 
        Event<PropertyChangedEventHandler,PropertyChangedEventArgs>()
    member this.Notify name =
        propertyChanged.Trigger(this,PropertyChangedEventArgs name)
    interface INotifyPropertyChanged with
        [<CLIEvent>]
        member this.PropertyChanged = propertyChanged.Publish

type DataGrid(headings:seq<_>, items:IEnumerable, cellFactory:int*int->FrameworkElement) as grid =
    inherit Grid()    
    do  grid.ShowGridLines <- true   
    let createHeader heading horizontalAlignment =
        let header = TextBlock(Text=heading)
        header.HorizontalAlignment <- horizontalAlignment
        header.VerticalAlignment <- VerticalAlignment.Center
        let container = Grid(Background=SolidColorBrush Colors.Gray)        
        container.Children.Add header |> ignore
        container
    do  ColumnDefinition(Width=GridLength(24.0)) |> grid.ColumnDefinitions.Add 
    do  headings |> Seq.iteri (fun i heading ->
        let width = GridLength(64.0)
        ColumnDefinition(Width=width) |> grid.ColumnDefinitions.Add       
        let header = createHeader heading HorizontalAlignment.Center
        grid.Children.Add header |> ignore
        Grid.SetColumn(header,i+1)
    )   
    do  let height = GridLength(24.0)
        RowDefinition(Height=height) |> grid.RowDefinitions.Add
        let mutable y = 1
        for item in items do
        RowDefinition(Height=height) |> grid.RowDefinitions.Add
        let header = createHeader (y.ToString()) HorizontalAlignment.Right       
        grid.Children.Add header |> ignore
        Grid.SetRow(header,y)
        for x=1 to Seq.length headings do
            let cell = cellFactory (x-1,y-1)
            cell.DataContext <- item
            grid.Children.Add cell |> ignore
            Grid.SetColumn(cell,x)
            Grid.SetRow(cell,y)
        y <- y + 1

type View() =
    inherit UserControl()  
    let sheet = Sheet(26,50)
    let remember = ignore
    let cellFactory (x,y) =
        let binding = Binding(sprintf "Cells.[%d].Data" x)   
        binding.Mode <- BindingMode.TwoWay
        let edit = TextBox()
        edit.SetBinding(TextBox.TextProperty,binding) |> ignore        
        edit.Visibility <- Visibility.Collapsed
        let view = Button(Background=SolidColorBrush Colors.White)
        view.BorderBrush <- null
        let binding = Binding(sprintf "Cells.[%d].Value" x)
        let block = TextBlock()
        block.SetBinding(TextBlock.TextProperty, binding) |> ignore
        view.Content <- block
        view.HorizontalContentAlignment <- HorizontalAlignment.Left
        view.VerticalContentAlignment <- VerticalAlignment.Center
        let setEditMode _ =
            edit.Visibility <- Visibility.Visible
            view.Visibility <- Visibility.Collapsed                   
            edit.Focus() |> ignore
        let setViewMode _ =
            edit.Visibility <- Visibility.Collapsed
            view.Visibility <- Visibility.Visible        
        view.Click |> Observable.subscribe setEditMode |> remember
        edit.LostFocus |> Observable.subscribe setViewMode |> remember        
        let enterKeyDown = edit.KeyDown |> Observable.filter (fun e -> e.Key = Key.Enter)
        enterKeyDown |> Observable.subscribe setViewMode |> remember
        let container = Grid()
        container.Children.Add view |> ignore
        container.Children.Add edit |> ignore
        container :> FrameworkElement
    let viewer = ScrollViewer(HorizontalScrollBarVisibility=ScrollBarVisibility.Auto)
    do  viewer.Content <- DataGrid(sheet.Columns,sheet.Rows,cellFactory)
    do  base.Content <- viewer

#if INTERACTIVE
open Microsoft.TryFSharp
App.Dispatch (fun() -> 
    App.Console.ClearCanvas()
    new View() |> App.Console.Canvas.Children.Add
    App.Console.CanvasPosition <- CanvasPosition.Right
)
#endif