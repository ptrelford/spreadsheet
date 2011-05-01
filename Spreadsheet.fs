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

type UnitType =
    | Unit of string * int 
    | CompositeUnit of UnitType list     
    override this.ToString() =
        let exponent = function
            | Unit(_,n) -> n
            | CompositeUnit(_) ->                
                raise (new System.InvalidOperationException())
        let rec toString = function        
            | Unit(s,n) when n=0 -> ""
            | Unit(s,n) when n=1 -> s
            | Unit(s,n)          -> s + " ^ " + n.ToString()            
            | CompositeUnit(us) ->               
                let ps, ns =
                    us |> List.partition (fun u -> exponent u >= 0)
                let join xs = 
                    let s = xs |> List.map toString |> List.toArray             
                    System.String.Join(" ",s)
                match ps,ns with
                | ps, [] -> join ps
                | ps, ns ->
                    let ns = ns |> List.map UnitType.Reciprocal
                    join ps + " / " + join ns
        match this with
        | Unit(_,n) when n < 0 -> " / " + toString this
        | _ -> toString this    
    static member ( * ) (v:ValueType,u:UnitType) = UnitValue(v,u)    
    static member ( * ) (lhs:UnitType,rhs:UnitType) =
        let text = function
            | Unit(s,n) -> s
            | CompositeUnit(us) -> us.ToString()
        let normalize us u =
            let t = text u
            match us |> List.tryFind (fun x -> text x = t), u with
            | Some(Unit(s,n) as v), Unit(_,n') ->
                us |> List.map (fun x -> if x = v then Unit(s,n+n') else x)                 
            | Some(_), _ -> raise (new System.NotImplementedException())
            | None, _ -> us@[u]
        let normalize' us us' =
            us' |> List.fold (fun (acc) x -> normalize acc x) us
        match lhs,rhs with
        | Unit(u1,p1), Unit(u2,p2) when u1 = u2 ->
            Unit(u1,p1+p2)       
        | Unit("",p1), _ -> rhs
        | _, Unit("",p2) -> lhs 
        | Unit(u1,p1), Unit(u2,p2) ->            
            CompositeUnit([lhs;rhs])
        | CompositeUnit(us), Unit(_,_) ->
            CompositeUnit(normalize us rhs)
        | Unit(_,_), CompositeUnit(us) ->
            CompositeUnit(normalize' [lhs]  us)
        | CompositeUnit(us), CompositeUnit(us') ->
            CompositeUnit(normalize' us us')
        | _,_ -> raise (new System.NotImplementedException())
    static member Reciprocal x =
        let rec reciprocal = function
            | Unit(s,n) -> Unit(s,-n)
            | CompositeUnit(us) -> CompositeUnit(us |> List.map reciprocal)
        reciprocal x
    static member ( / ) (lhs:UnitType,rhs:UnitType) =        
        lhs * (UnitType.Reciprocal rhs)
    static member ( + ) (lhs:UnitType,rhs:UnitType) =       
        if lhs = rhs then lhs                
        else raise (new System.InvalidOperationException())
and ValueType = decimal
and UnitValue(v:ValueType,u:UnitType) = 
    member this.Value = v 
    member this.Unit = u
    override this.ToString() = sprintf "%O %O" v u
    static member (~-) (v:UnitValue) =
        UnitValue(-v.Value,v.Unit)
    static member (+) (lhs:UnitValue,rhs:UnitValue) =
        UnitValue(lhs.Value+rhs.Value, lhs.Unit+rhs.Unit)         
    static member (-) (lhs:UnitValue,rhs:UnitValue) =
        UnitValue(lhs.Value-rhs.Value, lhs.Unit+rhs.Unit) 
    static member (*) (lhs:UnitValue,rhs:UnitValue) =                    
        UnitValue(lhs.Value*rhs.Value,lhs.Unit*rhs.Unit)                
    static member (*) (lhs:UnitValue,rhs:ValueType) =        
        UnitValue(lhs.Value*rhs,lhs.Unit)      
    static member (*) (v:UnitValue,u:UnitType) = 
        UnitValue(v.Value,v.Unit*u)  
    static member (/) (lhs:UnitValue,rhs:UnitValue) =                    
        UnitValue(lhs.Value/rhs.Value,lhs.Unit/rhs.Unit)
    static member (/) (lhs:UnitValue,rhs:ValueType) =
        UnitValue(lhs.Value/rhs,lhs.Unit)  
    static member (/) (v:UnitValue,u:UnitType) =
        UnitValue(v.Value,v.Unit/u)
    override this.Equals(that) =
        let that = that :?> UnitValue
        this.Unit = that.Unit && this.Value = that.Value
    override this.GetHashCode() = 
        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(this)
    interface System.IComparable with
        member this.CompareTo(that) =
            let that = that :?> UnitValue
            if this.Unit = that.Unit then
                if this.Value < that.Value then -1
                elif this.Value > that.Value then 1
                else 0
            else invalidOp "Unit mismatch"

module Unit =
    let Empty = Unit("",0)
    let Create s = Unit(s,1)
    let ValueOf(v,u) = UnitValue(v, Create(u))

type token =
    | WhiteSpace
    | Symbol of char
    | OpToken of string
    | RefToken of int * int
    | StrToken of string
    | NumToken of string

let (|Match|_|) pattern input =
    let m = System.Text.RegularExpressions.Regex.Match(input, pattern)
    if m.Success then Some m.Value else None

let rec toRef (s:string) =
    let col = int s.[0] - int 'A'
    let row = s.Substring 1 |> int
    col, row-1

let matchToken = function
    | Match @"^\s+" s -> s, WhiteSpace
    | Match @"^\+|^\-|^\*|^\/|^\^"  s -> s, OpToken s
    | Match @"^=|^<>|^<=|^>=|^>|^<"  s -> s, OpToken s   
    | Match @"^\(|^\)|^\,|^\:" s -> s, Symbol s.[0]   
    | Match @"^[A-Z]\d+" s -> s, s |> toRef |> RefToken
    | Match @"^[A-Za-z]+" s -> s, StrToken s
    | Match @"^\d+(\.\d+)?|\.\d+" s -> s, s |> NumToken
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

type arithmeticOp = Add | Sub | Mul | Div
type logicalOp = Eq | Lt | Gt | Le | Ge | Ne
type formula =
    | Neg of formula
    | ArithmeticOp of formula * arithmeticOp * formula
    | LogicalOp of formula * logicalOp * formula
    | Num of UnitValue
    | Ref of int * int
    | Range of int * int * int * int
    | Fun of string * formula list

let rec (|Term|_|) = function
    | Sum(f1, OpToken(LogicOp op)::Sum(f2,t)) -> Some(LogicalOp(f1,op,f2),t)
    | Sum(f1,t) -> Some (f1,t)
    | _ -> None
and (|LogicOp|_|) = function
    | "=" ->  Some Eq | "<>" -> Some Ne
    | "<" ->  Some Lt | ">"  -> Some Gt
    | "<=" -> Some Le | ">=" -> Some Ge
    | _ -> None
and (|Sum|_|) = function
    | Factor(f1, t) ->      
        let rec aux f1 = function        
            | SumOp op::Factor(f2, t) -> aux (ArithmeticOp(f1,op,f2)) t               
            | t -> Some(f1, t)      
        aux f1 t  
    | _ -> None
and (|SumOp|_|) = function 
    | OpToken "+" -> Some Add | OpToken "-" -> Some Sub 
    | _ -> None
and (|Factor|_|) = function  
    | OpToken "-"::Factor(f, t) -> Some(Neg f, t)
    | Atom(f1, ProductOp op::Factor(f2, t)) ->
        Some(ArithmeticOp(f1,op,f2), t)       
    | Atom(f, t) -> Some(f, t)  
    | _ -> None    
and (|ProductOp|_|) = function
    | OpToken "*" -> Some Mul | OpToken "/" -> Some Div
    | _ -> None
and (|Atom|_|) = function
    | NumToken n::Units(u,t) -> Some(Num(u * decimal n),t)
    | NumToken n::t -> Some(Num(UnitValue(decimal n,Unit.Empty)), t)
    | RefToken(x1,y1)::Symbol ':'::RefToken(x2,y2)::t -> 
        Some(Range(min x1 x2,min y1 y2,max x1 x2,max y1 y2),t)  
    | RefToken(x,y)::t -> Some(Ref(x,y), t)
    | Symbol '('::Term(f, Symbol ')'::t) -> Some(f, t)
    | StrToken s::Tuple(ps, t) -> Some(Fun(s,ps),t)  
    | Units(u,t) -> Some(Num(u),t)    
    | _ -> None
and (|Units|_|) = function
    | Unit'(u,t) ->
        let rec aux u1 =  function
            | OpToken "/"::Unit'(u2,t) -> aux (u1 / u2) t
            | Unit'(u2,t) -> aux (u1 * u2) t
            | t -> Some(u1,t)
        aux u t
    | _ -> None
and (|Unit'|_|) = function
    | StrToken u::OpToken "^"::NumToken p::t -> 
        Some(UnitValue(1.0M,Unit(u,int p)),t)
    | StrToken u::t ->
        Some(Unit.ValueOf(1.0M,u), t)
    | _ -> None
and (|Tuple|_|) = function
    | Symbol '('::Params(ps, Symbol ')'::t) -> Some(ps, t)  
    | _ -> None
and (|Params|_|) = function
    | Term(f1, t) ->
        let rec aux fs = function
            | Symbol ','::Term(f2, t) -> aux (fs@[f2]) t
            | t -> fs, t
        Some(aux [f1] t)
    | t -> Some ([],t)

let parse s = 
    tokenize s |> function 
    | Term(f,[]) -> f 
    | _ -> failwith "Failed to parse formula"

let evaluate valueAt formula =
    let rec eval = function
        | Neg f -> - (eval f)
        | ArithmeticOp(f1,op,f2) -> arithmetic op (eval f1) (eval f2)
        | LogicalOp(f1,op,f2) -> 
            if logic op (eval f1) (eval f2) 
            then UnitValue(0.0M,Unit.Empty) 
            else UnitValue(-1.0M,Unit.Empty)
        | Num d -> d
        | Ref(x,y) -> valueAt(x,y)
        | Range _ -> invalidOp "Range expected in function"
        | Fun("SUM",ps) -> ps |> evalAll |> List.reduce (+)
        | Fun("IF",[condition;f1;f2]) -> 
            if (eval condition).Value=0.0M then eval f1 else eval f2 
        | Fun(_,_) -> failwith "Unknown function"        
    and arithmetic = function
        | Add -> (+) | Sub -> (-) | Mul -> (*) | Div -> (/)
    and logic = function         
        | Eq -> (=)  | Ne -> (<>)
        | Lt -> (<)  | Gt -> (>)
        | Le -> (<=) | Ge -> (>=)
    and evalAll ps =
        ps |> List.collect (function            
            | Range(x1,y1,x2,y2) ->
                [for x=x1 to x2 do for y=y1 to y2 do yield valueAt(x,y)]
            | x -> [eval x]            
        )
    eval formula

let references formula =
    let rec traverse = function
        | Ref(x,y) -> [x,y]
        | Range(x1,y1,x2,y2) -> 
            [for x=x1 to x2 do for y=y1 to y2 do yield x,y]
        | Fun(_,ps) -> ps |> List.collect traverse
        | ArithmeticOp(f1,_,f2) | LogicalOp(f1,_,f2) -> 
            traverse f1 @ traverse f2
        | _ -> []
    traverse formula

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
    let mutable unitValue = UnitValue(0.0M,Unit("",0))
    let mutable data = ""       
    let mutable formula : formula option = None
    let updated = Event<_>()
    let mutable subscriptions : System.IDisposable list = []   
    let cellAt(x,y) = 
        let (row : Row) = Array.get sheet.Rows y
        let (cell : Cell) = Array.get row.Cells x
        cell
    let valueAt address = (cellAt address).UnitValue
    let eval formula =         
        try let v = evaluate valueAt formula in v, v.ToString()
        with _ -> UnitValue(0.0M,Unit("",0)), "N/A"
    let parseFormula (text:string) =
        if text.StartsWith "="
        then                
            try true, parse (text.Substring 1) |> Some
            with _ -> true, None
        else false, None
    let update (newUnitValue, newValue) generation =
        if newValue <> value || newUnitValue <> unitValue then
            value <- newValue  
            unitValue <- newUnitValue         
            updated.Trigger generation
            cell.Notify "Value"          
    let unsubscribe () =
        subscriptions |> List.iter (fun d -> d.Dispose())
        subscriptions <- []
    let subscribe formula addresses =
        let remember x = subscriptions <- x :: subscriptions
        for address in addresses do
            let cell' : Cell = cellAt address
            cell'.Updated
            |> Observable.subscribe (fun generation ->   
                if generation < sheet.MaxGeneration then
                    let newValue = eval formula
                    update newValue (generation+1)
            ) |> remember
    member cell.Data 
        with get () = data 
        and set (text:string) =
            data <- text        
            cell.Notify "Data"          
            let isFormula, newFormula = parseFormula text               
            formula <- newFormula
            unsubscribe()
            formula |> Option.iter (fun f -> references f |> subscribe f)
            let newValue =
                match isFormula, formula with           
                | _, Some f -> eval f
                | true, _ -> UnitValue(0.0M,Unit.Empty), "N/A"
                | _, None -> UnitValue(0.0M,Unit.Empty), text
            update newValue 0
    member cell.Value = value
    member cell.UnitValue = unitValue
    member cell.Updated = updated.Publish
and ObservableObject() =
    let propertyChanged = 
        Event<PropertyChangedEventHandler,PropertyChangedEventArgs>()
    member this.Notify name =
        propertyChanged.Trigger(this,PropertyChangedEventArgs name)
    interface INotifyPropertyChanged with
        [<CLIEvent>]
        member this.PropertyChanged = propertyChanged.Publish

let (+.) (grid:Grid) (child:#UIElement) = grid.Children.Add child |> ignore; grid

type DataGrid(headings:seq<_>, items:IEnumerable, cellFactory:int*int->FrameworkElement) as grid =
    inherit Grid(ShowGridLines=true)    
    let createHeader heading horizontalAlignment =
        let header = TextBlock(Text=heading)
        header.HorizontalAlignment <- horizontalAlignment
        header.VerticalAlignment <- VerticalAlignment.Center
        Grid(Background=SolidColorBrush Colors.Gray) +. header
    do  ColumnDefinition(Width=GridLength(24.0)) |> grid.ColumnDefinitions.Add 
    do  headings |> Seq.iteri (fun i heading ->
        let width = GridLength(128.0)
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
    let cellFactory (x,y) =
        let binding = Binding(sprintf "Cells.[%d].Data" x)   
        binding.Mode <- BindingMode.TwoWay
        let edit = TextBox()
        edit.SetBinding(TextBox.TextProperty,binding) |> ignore        
        edit.Visibility <- Visibility.Collapsed
        let block = TextBlock()
        let binding = Binding(sprintf "Cells.[%d].Value" x)
        block.SetBinding(TextBlock.TextProperty, binding) |> ignore
        let view = Button(Content=block)
        view.Background <- SolidColorBrush Colors.White
        view.BorderBrush <- null        
        view.HorizontalContentAlignment <- HorizontalAlignment.Left
        view.VerticalContentAlignment <- VerticalAlignment.Center
        let setEditMode _ =
            edit.Visibility <- Visibility.Visible
            view.Visibility <- Visibility.Collapsed                   
            edit.Focus() |> ignore
        let setViewMode _ =
            edit.Visibility <- Visibility.Collapsed
            view.Visibility <- Visibility.Visible        
        view.Click |> Observable.subscribe setEditMode |> ignore
        edit.LostFocus |> Observable.subscribe setViewMode |> ignore
        let enterKeyDown = edit.KeyDown |> Observable.filter (fun e -> e.Key = Key.Enter)
        enterKeyDown |> Observable.subscribe setViewMode |> ignore
        Grid() +. view +. edit :> FrameworkElement
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