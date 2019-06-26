# VBA Style Guide()

*A mostly reasonable approach to VBA.*

> Note: this is in it's **early stages and is currently under construction**. I will be adding to it as I can!

## Table of Contents
 
 1. [Naming Conventions](#naming-conventions)
 1. [Variables](#variables)
 1. [Functions](#functions)
 1. [Comments](#comments)
 1. [Performance](#performance)
 1. [Design](#design)

## Naming Conventions

  <a name="single--letter--names"></a><a name="1.1"></a>
  - [1.1](#single--letter--names) Avoid single letter names. Be descriptive with your naming.
    
    ```vb
    ' bad
    Function Q ()
        Dim i as Long
        ' ...
    End Function

    ' good
    Function Query ()
        Dim recordIndex as Long
        ' ...
    End Function
    ```

  <a name="pascal--case"></a><a name="1.2"></a>
  - [1.2](#pascal--case) Use PascalCase as the default naming convention for anything global.
    ```vb
    ' good
    Function GreetUser ()
        ' ...
    End Function
    ```
    
  <a name="camel--case"></a><a name="1.3"></a>
  - [1.3](#camel--case) Use camelCase for parameters and local variables and functions.
  
    > [Microsofts convention](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/naming-conventions) is **PascalCase** for everything. The most important thing is to be consistent in whatever convention you use.
    ```vb
    ' good
    Private Function sayName (ByVal name as string)
        ' ...
    End Function
    ```

  <a name="underscore--case"></a><a name="1.4"></a>
  - [1.4](#underscore--case) Do not use underscore case.
    
    > Why? VBA uses underscores for pointing out events and implementation. In fact, you can't implement another class if the other class has any public methods or properties with an underscore in the name otherwise you will get the error [Bad interface for Implements: method has underscore in name](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/bad-interface-for-implements-method-has-underscore-in-name).
    ```vb
    ' bad
    Dim First_Name as String

    ' good
    Dim firstName as String
    ```
    
  <a name="system--hungarian"></a><a name="1.5"></a>
  - [1.5](#system--hungarian) Do not use Systems Hungarian.
  
    > Why? These are useless prefixes that serve no purpose and can obscure the variables name.
    ```vb
    ' very bad
    Dim strString as String
    Dim oRange as Range

    ' bad
    Dim sName as String
    Dim rngData as Range
    Dim iCount as Integer

    ' good
    Dim firstName as String
    Dim queryData as Range
    Dim rowIndex as Integer
    ```

  <a name="abbreviations"></a><a name="1.6"></a>
  - [1.6](#abbreviations) Do not use abbreviations.
    ```vb
    ' bad
    Function GetWin()
        ' ...
    End Function

    ' good
    Function GetWindow()
        ' ...
    End Function
    ```

  <a name="descriptive--names"></a><a name="1.7"></a>
  - [1.7](#descriptive--names) Be descriptive and use easily readable identifier names. Programming is more about reading code!
    ```vb
    ' very bad
    Dim x As Boolean

    ' bad
    Dim scrollableX As Boolean

    ' good
    Dim canScrollHorizontally As Boolean
    ```

  **[⬆ back to top](#table-of-contents)**

## Variables

  <a name="declare-variables-where-used"></a><a name="2.1"></a>
  - [2.1](#declare-variables-where-used") Declare and assign variables next to where they are going to be used, but place them in a reasonable place.

  > Why? This makes maintaining the code much easier. When you have a wall of declarations at the top of a procedure it is difficult to modify and refactor if needed. Also, you have to scroll up and down to see if a variable is used or not.
  ```vb
    ' bad
    Private Sub someMethod(ByVal path As String)
    
        Dim fileSystem As Object
        Dim folder As Object
        Dim files As Object
        Dim file As Object
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set folder = FSO.GetFolder(path)
        Set files = folder.Files

        For Each file In files
            '...
        Next

    End Sub

    ' good
    Private Sub someMethod(ByVal path As String)
    
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        Dim folder As Object
        Set folder = FSO.GetFolder(path)
        
        Dim files As Object
        Set files = folder.Files
        
        Dim file As Object
        For Each file In files
            '...
        Next

    End Sub
  ```

  <a name="keep-variables-local"></a><a name="2.2"></a>
  - [2.2](#keep-variables-local) Prefer to keep variables local using the `Private` keyword. We want to avoid polluting the global namespace. Captain Planet warned us of that.
      ```vb
    ' bad
    Public Const FileName as string = "C:\"

    ' good
    Private Const fileName as string = "C:\"
    ```

  <a name="no-unused-variables"></a><a name="2.3"></a>
  - [2.3](#no-unused-variables) Disallow unused variables.

    > Why? Variables that are declared and not used anywhere in the code are most likely an error due to incomplete refactoring. Such variables take up space in the code and can lead to confusion by readers.

    ```vb
    ' bad
    Dim someUnusedVariable as String

    ' good
    Dim message as string
    message = "I will be used!"
    Msgbox Messgae
    ```

  <a name="use-option-explicit"></a><a name="2.4"></a>
  - [2.4](#use-option-explicit) Use [`Option Explicit`](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement) to ensure all variables are explicitly declared.
  
    ```vb
    ' good
    Option Explicit

    Sub doSomething()
        x = 1 ' <~ Compile error: Variable not defined
    End Sub
    ```

  <a name="no-one-line-declarations"></a><a name="2.5"></a>
  - [2.5](#no-one-line-declarations) Use one `Dim` declaration per variable or assignment.

    > Why? It's easier to read and debug going back. It also prevents variables from accidentally being declared as Variants.
  
    ```vb
    ' very bad
    Dim lastRow, lastColumn As Long '<~ lastRow is a Variant, NOT a long

    ' bad
    Dim lastRow As Long, lastColumn As Long

    ' good
    Dim lastRow As Long
    Dim lastColumn As Long
    ```

  <a name="decalare-variable-types"></a><a name="2.6"></a>
  - [2.6](#decalare-variable-types) Declare all variable types explicitly.

    ```vb
    ' bad
    Dim row
    Dim name
    Dim cell

    ' good
    Dim row As Long
    Dim name As String
    Dim cell As Range
    ```

  **[⬆ back to top](#table-of-contents)**

## Functions

  <a name="functions--mutate-params"></a><a name="3.1"></a>
  - [3.1](#functions--mutate-params") Prefer `ByVal` for parameters.

    > Why? Reassigning and mutating parameters can lead to unexpected behavior and errors. `ByRef` is very helpful at times, but the general rule is to default to `ByVal`.

    ```vb
    ' bad
    Function doSomething(name As String) As String

    ' ok
    Function doSomething(ByRef outName As String) As Boolean

    ' good
    Function doSomething(ByVal name As String) As String
    ```


## Comments

  <a name="description-header-comment"></a><a name="4.1"></a>
  - [4.1](#description-header-comment) Above the function should be a simple description of what the function does. Keep it simple.
    ```vb
    ' Adds new element(s) to an array (at the end) and returns the new array length.
    Function PushToArray(ByRef SourceArray As Variant, ParamArray Element() As Variant) As Long
       '...
    End Function
    ```
  

  <a name="doc--comment"></a><a name="4.2"></a>
  - [4.2](#doc--comment) Just inside the function is where I will put important details. This could be the author, library references, notes, Ect. I've styled this to be similar to [JSDoc documentation](https://devdocs.io/jsdoc/). 
     ```vb
    ' Adds new element(s) to an array (at the end) and returns the new array length.
    Function PushToArray(ByRef SourceArray As Variant, ParamArray Element() As Variant) As Long
        ' @author: Robert Todar <https://github.com/todar>
        ' @param: <SourceArray> can be either 1 or 2 dimensional array.
        ' @param: <Element> are the elements to be added.
        ' @ref: No Library references needed =)
        
        '...
    End Function
    ```

  <a name="descriptive--comment"></a><a name="4.3"></a>
  - [4.3](#descriptive--comment) Notes should be clear and full sentences. Explain anything that doesn't immediately make sense from the code.
    ```vb
    'Need to check to make sure there are records to pull from.
    If rs.BOF Or rs.EOF Then
        Exit Function
    End If
    ```

  <a name="actionitems"></a><a name="4.4"></a>
  - [4.4](#actionitems) Prefixing your comments with `FIXME` or `TODO` helps other developers quickly understand if you’re pointing out a problem that needs to be revisited, or if you’re suggesting a solution to the problem that needs to be implemented. These are different than regular comments because they are actionable. The actions are `FIXME: -- need to figure this out` or `TODO: -- need to implement`.

  **[⬆ back to top](#table-of-contents)**

## Performance

  <a name="avoid-using-select"></a><a name="5.1"></a>
  - [5.1](#avoid-using-select) [Avoid using select in Excel](https://stackoverflow.com/q/10714251/8309643).
   
    > Why? It slows down code and also can cause runtime errors. `Select` should only be used for visual reasons such as the users next task is doing something in that specific cell.
    
    ```vb
    ' bad
    Range("A1").Select
    ActiveCell.Value = "Darth Vader"
    
    ' ok
    Dim cell As Range
    Set cell = ActiveSheet.ActiveCell
    cell.Value = "Lando Calrissian"
    
    ' good
    With Workbooks("Star Wars").Worksheets("Characters").Range("Hero")
        .Value = "Luke Skywalker"
    End With
    ```

  **[⬆ back to top](#table-of-contents)**

## Design

  - Functions should be small.
  - Functions should be pure.
  - Anytime there is a section of code that is separated by a giant comment block, ask yourself if this needs to get extracted into it's own function.

  **[⬆ back to top](#table-of-contents)**
