# VBA Style Guide()

*A mostly reasonable approach to VBA.*

> Note: this is in it's **early stages and is currently under contruction**. I will be adding to it as I can!

## Table of Contents
 
 1. [Naming Conventions](#naming-conventions)
 1. [Variables](#variables)
 1. [Comments](#comments)
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
      Dim RecordIndex as Long
      ' ...
    End Function
    ```

  <a name="pascal--case"></a><a name="1.2"></a>
  - [1.2](#pascal--case) Use PascalCase as the default naming convention.
    ```vb
    ' good
    Function GreetUser ()
      ' ...
    End Function
    ```
    
  <a name="camel--case"></a><a name="1.3"></a>
  - [1.3](#camel--case) Use camelCase for parameters and local variables and functions.
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
    Dim FirstName as String
    ```
  **[⬆ back to top](#table-of-contents)**

## Variables

  <a name="declare-where-used"></a><a name="2.1"></a>
  - [2.1](#declare-where-used) Declare variables next to where they are going to be used.

  > Why? This makes maintaing the code much easier. When you have a wall of declarations at the top of a procedure it is difficult modify and refactor if needed. Also, you have to scroll up and down to see if a variable is used or not.
  ```vb
    ' bad
    Private Sub SomeMethod(ByVal Path As String)
    
        Dim FSO As Object
        Dim Folder As Object
        Dim Files As Object
        Dim File As Object
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set Folder = FSO.GetFolder(Path)
        Set Files = Folder.Files

        For Each File In Files
            '...
        Next

    End Sub

    ' good
    Private Sub SomeMethod(ByVal Path As String)
    
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        Dim Folder As Object
        Set Folder = FSO.GetFolder(Path)
        
        Dim Files As Object
        Set Files = Folder.Files
        
        Dim File As Object
        For Each File In Files
            '...
        Next

    End Sub
  ```

  <a name="keep-variables-local"></a><a name="2.2"></a>
  - [2.2](#keep-variables-local) Aim to keep variables local using the `Private` keyword. We want to avoid polluting the global namespace. Captain Planet warned us of that.
      ```vb
    ' bad
    Public Const FileName as string = "C:\"

    ' good
    Private Const FileName as string = "C:\"
    ```
  **[⬆ back to top](#table-of-contents)**

## Comments

  <a name="description-header-comment"></a><a name="3.1"></a>
  - [3.1](#description-header-comment) Above the function should be a simple description of what the function does.

  <a name="doc--comment"></a><a name="3.2"></a>
  - [3.1](#doc--comment) Just inside the function is where I will put important details. This could be author, library references, notes, Ect. I've styled this to be similar to [JSDoc documentation](https://devdocs.io/jsdoc/). 

  <a name="descriptive--comment"></a><a name="3.1"></a>
  - [3.1](#descriptive--comment) Notes should be clear and full sentences. Explain anything that doesn't immediatly make sence from the code.

  **[⬆ back to top](#table-of-contents)**


## Design

  Functions should be as small as possible designed to resusable. This means they should be very readable. There should not be giant procedures. Anytime there is a section of code that is seperated by a giant comment block, ask yourself if this needs to get extracted into it's own function.

  **[⬆ back to top](#table-of-contents)**
