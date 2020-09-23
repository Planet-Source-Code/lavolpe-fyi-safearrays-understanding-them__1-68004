VERSION 5.00
Begin VB.Form frmArrayPtrs 
   Caption         =   "Click Button to Get Array Information"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Null Array"
      Height          =   495
      Index           =   9
      Left            =   5820
      TabIndex        =   10
      Top             =   690
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Array in UDT"
      Height          =   495
      Index           =   8
      Left            =   4485
      TabIndex        =   9
      Top             =   690
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dates"
      Height          =   495
      Index           =   7
      Left            =   3120
      TabIndex        =   8
      Top             =   690
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Longs"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bytes"
      Height          =   495
      Index           =   1
      Left            =   1740
      TabIndex        =   6
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Variants"
      Height          =   495
      Index           =   2
      Left            =   3090
      TabIndex        =   5
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "stdPictures"
      Height          =   495
      Index           =   3
      Left            =   4455
      TabIndex        =   4
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VarLen Strings"
      Height          =   495
      Index           =   4
      Left            =   5775
      TabIndex        =   3
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UDTs"
      Height          =   495
      Index           =   6
      Left            =   1755
      TabIndex        =   2
      Top             =   690
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Static Strings"
      Height          =   495
      Index           =   5
      Left            =   375
      TabIndex        =   1
      Top             =   690
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3495
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmArrayPtrs.frx":0000
      Top             =   1335
      Width           =   7110
   End
End
Attribute VB_Name = "frmArrayPtrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' References:
' SafeArray structure: http://msdn2.microsoft.com/en-us/library/ms892133.aspx
' SafeArray Bounds structure: http://msdn2.microsoft.com/en-us/library/ms892134.aspx
' SafeArray .fFeatures values: http://msdn2.microsoft.com/en-us/library/ms221482.aspx
' SafeArray VarTypes: http://msdn2.microsoft.com/en-us/library/ms891853.aspx

' A little fun with SafeArrays; how to read them, how to get them, and how to create/use them

' With VB5 & VB6 (VB4?) we can determine the structure, dimensions & other information of an array
' by passing it into a Variant and then parsing the variant structure as has been posted many times
' on PSC. ULLI (an excellent coder) posted something similar to this routine and is probably the
' best of the bunch. However none of them can show you how to get all the array information from
' just a pointer and very few try to explain how to create and use SafeArrays.

' Previous version of this tutorial project was using a method to determine if an array
' is initialized or not that did not use the VarPtrArray API call. However, recent discoveries
' in another project seems to prove that that method can cause "Expression Too Complex" errors
' in the routine that called that function. What is more frustrating is that the error occurs
' intermittently and may/may not occur on different pcs using the same O/S, Therefore, the
' IsArrayEmpty routine has been modified and all lines of code that call that function now
' use VarPtrArray.

' Have fun playing with it. But beware, this is all strictly memory reads, and if you modify
' code or add your own, be prepared to crash if you are not very careful.

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
' VB5: Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Type SafeArrayBound
    cElements As Long           ' nr of array elements in this dimensions
    lLbound As Long             ' the lBound of this dimension. UBound = cElements + lLBound - 1
End Type
Private Type SafeArrayReader    ' Not usable to apply SafeArrays because the rgSABound cannot be dynamic
    cDims As Integer            ' nr of dimensions for the array
    fFeatures As Integer        ' extra information about the array contents
    cbElements As Long          ' nr of bytes per array element. Possible Examples: 1=byte,2=integer,4=long,8=currency
    cLocks As Long              ' nr of times array was locked w/o being unlocked
    pvData As Long              ' address to 1st array item, can be a pointer to another structure/address
    rgSABound() As SafeArrayBound
End Type


' Now for some added info on applying SafeArrays to overlay memory pointers in order
' to read data more easily.
' ---------------------------------------------------------------------------------------
' The following comments are aimed at understanding how to overlay SafeArrays onto memory
' ---------------------------------------------------------------------------------------

Private Type SafeArrayDynamic
    cDims As Integer            ' nr of dimensions for the overlay
    fFeatures As Integer        ' not needed for overlaying SafeArrays
    cbElements As Long          ' nr of bytes per array element. Examples: 1=byte,2=integer,4=long,8=currency
    cLocks As Long              ' not needed for overlaying SafeArrays
    pvData As Long              ' address of 1st array item
    rgSABound(0 To 2) As SafeArrayBound ' reusable for multiple dimensions. Make as large as you need
End Type

' ^^ About the rgSABound structure.
' These are in upside down / right to left order from the way we read them within
' a DIM statement. For example: myArray(1 to 10, 20 to 30) would have 2 rgSABound
' structures, but the order would be: rgSABound(0).lLBound=20 & rgSABound(1).lLBound=1
' -- Additionally, SumOf(rgSABound.cElements)*cbElements => bytes to be read/accessed

' ^^ About the pvData member.
' 1. If you know the memory address, then simply assign it: i.e., .pvData = dibPointer
' 2. If you want to overlay on an existing array: .pvData = VarPtr(existingArray(LBound(existingArray)))
' 3. If you want to overlay on a dynamic array, but can't use VarPtr because you don't know its LBound(s)
'   a. Read 1st 16 bytes of the passed array's SafeArray:
'       CopyMemory lFarPointer, ByVal VarPtrArray(passedArray), 4& ' lFarPointer is a Long variable
'       ensure lFarPointer is not zero, otherwise the passed array is null/not initialized & next line will crash
'       Copymemory mySafeArray, ByVal lFarPointer, 16&
'   b. Now the mySafeArray has the .pvData filled in for you.
'      Simply fill in the rest and clear any possible members that you won't use
'       mySafeArray.fFeatures=0
'       mySafeArray.cLocks=0
'       set .cDims, cbElements & rgSABound members as you need. See below
'   When would you not know the LBound(s) of any array? When that dynamic array is passed as a routine's parameter

' Ok, then why would you overlay a SafeArray on an array that you have full access to?
' Good question. Here is one example from my c32bppDIB Suite project.
' - That project's SetDIBbits routine accepts an array of bytes in any dimension, any LBound(s)
' - The routine needs to transfer those bytes onto a DIB where I do have a valid memory pointer
' - To make things easier to code, I want my DIB array as 2D, zero bound in both dimensions
' - And to prevent complicated offsets needed if my array and the passed array have different
'   LBounds and/or different dimensions, I overlay a 2D array onto the passed array. Now both
'   arrays have similar structures and looping thru the arrays are sooooo much simpler.


'Example of overlaying: Let's say you want to view 32bpp DIB data as bytes.

' I will pretend to have a DIB pointer for the example: dibPointer

' 1. Declare an uninitialized array of bytes: Dim dibBytes() As Byte
' Note. You will notice we never use ReDim. The SafeArray acts like a ReDim statement, however,
'       no memory is actually allocated/copied to the overlay. We are simply redirecting

' 2. Prepare your SafeArray variable: Dim tSA As SafeArrayDynamic
'   With tSA
'       .cbElements = 1                         ' bytes
'       .cDims = 2                              ' we'll view as 2D vs 1D
'       .pvData = dibPointer                    ' data memory address
'       .rgSABound(0).cElements = dibHeight
'       .rgSABound(0).lLBound = whatever you want, zero in our example
'       .rgSABound(1).cElements = dibWidth * 4  ' 32bpp DIB has 4 bytes per pixel
'       .rgSABound(1).lLBound = whatever you want, zero in our example
'   End With

' 3. Now we must use VarPtrArray() API call to get an address which really
' is a Pointer to the array's SafeArray structure. We don't use an already
' dimensioned/initialized array and overwrite its SafeArray pointer because
' you can easily crash or would need to restore the original pointer. The
' uninitialized SafeArray address is always zero.
'   CopyMemory ByVal VarPtrArray(dibBytes), VarPtr(tSA), 4&
'   Now the dibBytes would look like: dibBytes(0 to dibWidth-1, 0 to dibHeight-1)

' 4. After you are done needing the overlay, you must remove it. In IDE, and if
' not done, no side effects. But when compiled, you will crash if you don't
' remove the overlay.
'   CopyMemory ByVal VarPtrArray(dibBytes), 0&, 4&
' ---------------------------------------------------------------------------------------
' To view the above example as Longs vs Bytes, just tweak a couple of lines
' 1. Declare dibBytes as Long vs Bytes
' 2. Change .cbElements=4
' 3. Change .rgSABound(1).cElememts=dibWidth
' In fact you can setup up multiple SafeArrays overlaying on the same memory address
'     as long as you use uninitialized arrays for the overlays

' ---------------------------------------------------------------------------------------
' Although you can get tricky and simply navigate the DIB pointer using statements
' similar to the following... It is risky. I have included a SafeOffset function that
' can be used with CopyMemory.
'   For X=0 To totalDIBbytes-1 Step 4
'       CopyMemory lColor, ByVal dibPointer+X, 4&

' What's the risk? One: you are assuming the dibPointer is a positive value memory
' address. Two: what if the dibPointer is positive but the last byte is negative?
' Memory addresses are unsigned, but VB longs are signed so when 1 is added to
' the max positive value (bit speaking) the value becomes negative in VB. Now if
' someone can prove that VB will never use negative memory addresses, then this is n/a

' Want some added info on the SafeArray structure?
' Most of us are familiar with the TYPE declaration above, but
' in reality the SAFEARRAY structure is formatted more like the
' following where the VTinfo may or may not exist and can be 4 or 16
' bytes. Just know that VB's VarPtrArray API always returns a
' pointer to the .cDims address not the .cDims address itself

'    Private Type SafeArrayHeader   ' Not a real UDT, used to visualize only
'       VTinfo(0 To 15) As Byte     ' Depending on fFeatures. See comments below
'       SA As SafeArray
'    End Type
'    Private Type SafeArray
'        cDims As Integer
'        fFeatures As Integer       ' determines if extra info is avaiable
'        cbElements As Long
'        cLocks As Long
'        pvData As Long
'        rgSABound(0 To n) As SafeArrayBound
'    End Type

' - when fFeatures include &H200 or &H400 then there are 16 bytes of information
'       in front of the cDims member. Those 16 bytes are a GUID (i.e., stdPicture GUID).
' - when fFeatures include &H80 then there are 4 bytes of information
'       Those 4 bytes are the VarType of the array data (i.e., vbLong, vbDate, etc)
' - when fFeatures include &H20 then there are 4 bytes of information
'      Those 4 bytes are the pointer to an IRecordInfo interface
' - If fFeatures does not include &H20,&H40,&H200,&H400 then there should be no
'       extra information in front of the cDims and trying to access that memory
'       could cause a crash due to Access Violations



' //////////////////////////////////////////
' Following are only used for examples only
' //////////////////////////////////////////
Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String  ' 4 bytes (StrPtr) / pointer to String
    lpszClassName As String ' 4 bytes (StrPtr) / pointer to String
    hIconSm As Long
End Type

Private Const OFS_MAXPATHNAME As Long = 128
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type


Private Function IsArrayEmpty(ByVal FarPointer As Long, Optional ByRef ArrayPointer As Long) As Boolean
  ' Usage: IsArrayEmpty(VarPtrArray(myArray))
  CopyMemory ArrayPointer, ByVal FarPointer, 4&
  IsArrayEmpty = (ArrayPointer = 0)

End Function

Private Sub Command1_Click(Index As Integer)
    
    ' Sample Arrays
    Dim Xyz(1 To 5, 8 To 25) As Long
    Dim Abc() As Byte, Ddates() As Date
    Dim vArray(-20 To 6, 1962 To 2007, 3 To 9) As Variant
    Dim stdPics() As StdPicture
    Dim vStrings(10 To 20, -20 To -10) As String
    Dim sStrings(2001 To 2112) As String * 5
    Dim wcUDT(1 To 2) As WNDCLASSEX
    Dim ofsUDT As OFSTRUCT
    
    
    ' Safe Array structure and pointer we will need
    Dim tSA As SafeArrayReader, arrPointer As Long
    
    ' Formatting variables
    Dim sArrayStruct As String, sVarType As String
    Dim sOther As String, LB As Long
    
    Select Case Index
    Case 0
        If Not IsArrayEmpty(VarPtrArray(Xyz), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 2" & vbCrLf & _
                "Array Elements: Long" & vbCrLf & _
                "Pointer to first element: " & VarPtr(Xyz(1, 8)) & vbCrLf & _
                "Array Structure: (1 To 5, 8 To 25)" & vbCrLf & _
                "Other: Array is Static" & vbCrLf & vbCrLf
        End If
        
    Case 1
        LB = CInt(Rnd * 500)
        ReDim Abc(LB To LB + CInt(Rnd * 100))
        If Not IsArrayEmpty(VarPtrArray(Abc), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 1" & vbCrLf & _
                "Array Elements: Bytes" & vbCrLf & _
                "Pointer to first element: " & VarPtr(Abc(LB)) & vbCrLf & _
                "Array Structure: (" & LB & " To " & UBound(Abc) & ")" & vbCrLf & _
                "Other: Array is Dynamic" & vbCrLf & vbCrLf
        End If
    
    Case 2
        If Not IsArrayEmpty(VarPtrArray(vArray), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 3" & vbCrLf & _
                "Array Elements: Variant" & vbCrLf & _
                "Pointer to first element: " & VarPtr(vArray(-20, 1962, 3)) & vbCrLf & _
                "Array Structure: (-20 To 6, 1962 To 2007, 3 To 9)" & vbCrLf & _
                "Other: Array is Static" & vbCrLf & vbCrLf
        End If
        
    Case 3
        ReDim stdPics(1 To 10)
        If Not IsArrayEmpty(VarPtrArray(stdPics), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 1" & vbCrLf & _
                "Array Elements: stdPicture Objects" & vbCrLf & _
                "Pointer to first element: " & VarPtr(stdPics(1)) & vbCrLf & _
                "Array Structure: (1 To 10)" & vbCrLf & _
                "Other: Array is Dynamic" & vbCrLf & vbCrLf
        End If
        
    Case 4
        If Not IsArrayEmpty(VarPtrArray(vStrings), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 2" & vbCrLf & _
                "Array Elements: Variable Length Strings" & vbCrLf & _
                "Pointer to first element: " & VarPtr(vStrings(10, -20)) & vbCrLf & _
                "Array Structure: (10 To 20, -20 To -10)" & vbCrLf & _
                "Other: Array is Static" & vbCrLf & vbCrLf
        End If
        
    Case 5
        If Not IsArrayEmpty(VarPtrArray(sStrings), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 1" & vbCrLf & _
                "Array Elements: Static Length Strings - 5 Characters per string" & vbCrLf & _
                "Pointer to first element: " & StrPtr(sStrings(2001)) & vbCrLf & _
                "Array Structure: (2001 To 2112)" & vbCrLf & _
                "Other: Array is Static - SEE COMMENTS in routine regarding Static Length Strings" & vbCrLf & vbCrLf
            
            ' Note about fixed length strings: Something is going on with VB here.
            ' StrPtr(fixed_String_ArrayItem) does point to the string as validated by using CopyMemory
            ' the SafeArray's pvData member also points to the string as validated by using CopyMemory
            ' However, StrPtr() & pvData are different memory addresses. It appears VB makes a copy and
            ' is adding its own flag in the VT (varType) member associated with the SafeArray. You can
            ' see this below where the VT is being checked and the value of 32 is being returned. Per
            ' documentation I've found, 32 is not a valid VT value.
        End If
        
    Case 6
        If Not IsArrayEmpty(VarPtrArray(wcUDT), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 1" & vbCrLf & _
                "Array Elements: WNDCLASSEX udt (48 bytes)" & vbCrLf & _
                "Pointer to first element: " & VarPtr(wcUDT(1)) & vbCrLf & _
                "Array Structure: (1 To 2)" & vbCrLf & _
                "Other: Array is Static" & vbCrLf & vbCrLf
        End If
        
    Case 7
        LB = CInt(Rnd * 500)
        ReDim Ddates(LB To LB + CInt(Rnd * 100))
        If Not IsArrayEmpty(VarPtrArray(Ddates), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 1" & vbCrLf & _
                "Array Elements: Dates" & vbCrLf & _
                "Pointer to first element: " & VarPtr(Ddates(LB)) & vbCrLf & _
                "Array Structure: (" & LB & " To " & UBound(Ddates) & ")" & vbCrLf & _
                "Other: Array is Dynamic" & vbCrLf & vbCrLf
        End If
        
    Case 8
        If Not IsArrayEmpty(VarPtrArray(ofsUDT.szPathName), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: 1" & vbCrLf & _
                "Array Elements: OFSTRUCT udt .szPathName Member (Byte Array)" & vbCrLf & _
                "Pointer to first element: " & VarPtr(ofsUDT.szPathName(0)) & vbCrLf & _
                "Array Structure: (0 To " & UBound(ofsUDT.szPathName) & ")" & vbCrLf & _
                "Other: Array is Static and is part of the OFSTRUCT udt" & vbCrLf & vbCrLf
        End If
        
    Case 9
        If IsArrayEmpty(VarPtrArray(Abc), arrPointer) Then
            Text1.Text = "Hard Coded array property values:" & vbCrLf
            Text1.Text = Text1.Text & "Array Dimensions: none - not initialized" & vbCrLf & _
                "Array Elements: Bytes" & vbCrLf & _
                "Pointer to first element: n/a" & vbCrLf & _
                "Array Structure: n/a" & vbCrLf & _
                "Other: Array is Dynamic" & vbCrLf & vbCrLf
        End If
    
    End Select
        
        
        If arrPointer = 0 Then
        
            Text1.Text = Text1.Text & "Property values from memory addresses: " & vbCrLf & _
                "Array Elements: Unknown" & vbCrLf & _
                "Array Dimensions = not initialized" & vbCrLf & _
                "Pointer to first element: 0" & vbCrLf & _
                "Array Structure: not initialized" & vbCrLf & "Other: unknown"
            Exit Sub
        End If
        
        
        CopyMemory tSA, ByVal arrPointer, 16    ' copy up to SafeArray.rgSABound
        ReDim tSA.rgSABound(0 To tSA.cDims - 1) ' size our SafeArray Bounds
        ' now copy the bounds to our array
        CopyMemory tSA.rgSABound(0).cElements, ByVal SafeOffset(arrPointer, 16), tSA.cDims * 8
        ' ^ note. to be truly safe we will validate adding 16 won't overflow
        ' The SafeOffset function will return zero if the offset results in unsigned < 1 or unsigned > 2^32-1
        ' But since we are reading already allocated memory we won't make that check. We know result will be valid
        
        ' Start displaying what we are finding
        Text1.Text = Text1.Text & "Property values from memory addresses: " & vbCrLf & "Array Dimensions = " & tSA.cDims
        
        ' default array data types
        Select Case tSA.cbElements
            Case 1:     sVarType = "Bytes"
            Case 2:     sVarType = "Integer/Boolean/Char"
            Case 4:     sVarType = "Longs (Array Element Pointers)"
            Case 8:     sVarType = "Double/Currency/Date/Char*4"
            Case 16:    sVarType = "Variant/Char*8"
            Case Else:  sVarType = "Other (" & tSA.cbElements & " bytes per element)"
        End Select
            
        If (tSA.fFeatures And 128) = 128 Then ' we have the specific type of variable. See flag meanings below
            CopyMemory LB, ByVal SafeOffset(arrPointer, -4), 4&
            ' ^ note. to be truly safe we will validate subtracting 4 won't overflow
            ' The SafeOffset function will return zero if the offset results in unsigned < 1 or unsigned > 2^32-1
            ' But since we are reading already allocated memory we won't make that check. We know result will be valid
            Select Case LB
                Case vbByte:                sVarType = "Bytes"
                Case vbInteger:             sVarType = "Integers"
                Case vbBoolean:             sVarType = "Booleans"
                Case vbLong:                sVarType = "Longs"
                Case vbSingle:              sVarType = "Singles"
                Case vbDouble:              sVarType = "Doubles"
                Case vbDate:                sVarType = "Dates"
                Case vbCurrency:            sVarType = "Currencys"
                Case vbVariant:             sVarType = "Variants"
                Case 32:                    sVarType = "Static Length String * " & tSA.cbElements \ 2 & ", " & tSA.cbElements & " bytes per character)"
                Case vbUserDefinedType:     sVarType = "UDTs"
                Case vbDecimal:             sVarType = "Decimals"
                Case Else ' use the default from "Select Case tSA.cbElements" above
                
                '^^ about vbUserDefinedType. This will probably never happen within VB proper.
                '       The only time you will get this is an array of Public UDTs exportable from a referenced OCX or DLL
                '^^ about vbDecimal. What is this? It is equivalent to 2^96-1, 12 byte variable I believe
                
            End Select
        End If
        Text1.Text = Text1.Text & vbCrLf & "Array Elements: " & sVarType
        
        Text1.Text = Text1.Text & vbCrLf & "Pointer to first element: " & tSA.pvData
        
        For tSA.cDims = 0 To tSA.cDims - 1
            If Not tSA.cDims = 0 Then sArrayStruct = ", " & sArrayStruct
            LB = tSA.rgSABound(tSA.cDims).lLbound
            sArrayStruct = LB & " To " & tSA.rgSABound(tSA.cDims).cElements + LB - 1 & sArrayStruct
        Next
        Text1.Text = Text1.Text & vbCrLf & "Array Structure: (" & sArrayStruct & ")" & vbCrLf
        
        ' Some good information is available in the fFeatures member if you care to test it
        If (tSA.fFeatures And 2) = 0 Then sOther = "[Array is Dynamic] "
        LB = 2048
        Do
            Select Case (tSA.fFeatures And LB)
            Case 2048: sOther = sOther & "[An array of Variants] "
            Case 1024: sOther = sOther & "[An array of Objects] "
            Case 512: sOther = sOther & "[An array of IUnknowns] "
            Case 256: sOther = sOther & "[An array of Strings Pointers] " ' BSTRs
            Case 128: ' this is a flag indicating that 4 bytes before the safeArray header
                      ' we can find a value indicating specifically the element's data type
            Case 64: sOther = sOther & "[Elements have GUID Ref] "
                    ' note that we could get the 16 byte GUID if we needed it
            Case 32: sOther = sOther & "[Array of data records] "
            Case 16: sOther = sOther & "[Array cannot be resized] "
            Case 4: sOther = sOther & "[Array is part of a UDT] "
            Case 2: sOther = sOther & "[Array is Static] "
            Case 1: sOther = sOther & "[Array is stack allocated] "
            End Select
            LB = LB \ 2
        Loop Until LB = 0
        
        Text1.Text = Text1.Text & "Other: " & sOther
End Sub

Private Function SafeOffset(ByVal Ptr As Long, Offset As Long) As Long

    ' ref http://support.microsoft.com/kb/q189323/ ' unsigned math
    ' Purpose: Provide a valid/safe pointer offset. Primarily for use with CopyMemory
    
    ' If a pointer +/- the offset wraps around the high bit of a long, the
    ' pointer needs to change from positive to negative or vice versa.
    
    ' A return of zero indicates the offset exceeds the min/max unsigned long bounds
    ' Zero is an invalid memory address
    
    Const MAXINT_4NEG As Long = -2147483648#    ' min value of a Long
    Const MAXINT_4POS As Long = 2147483647      ' max value of a Long

    If Not Ptr = 0& Then ' zero is an invalid pointer
        If Offset < 0& Then ' subtracting from pointer
            If Ptr < MAXINT_4NEG - Offset Then
                ' wraps around high bit (backwards) & changes to Positive from Negative
                SafeOffset = MAXINT_4POS - ((MAXINT_4NEG - Ptr) - Offset - 1&)
            ElseIf Ptr > 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr > -Offset Then SafeOffset = Ptr + Offset
            Else
                SafeOffset = Ptr + Offset
            End If
        Else    ' Adding to pointer
            If Ptr > MAXINT_4POS - Offset Then
                ' wraps around high bit (forward) & changes to Negative from Positive
                SafeOffset = MAXINT_4NEG + (Offset - (MAXINT_4POS - Ptr) - 1&)
            ElseIf Ptr < 0& Then ' verify pointer does not wrap around 0 bit
                If Ptr < -Offset Then SafeOffset = Ptr + Offset
            Else
                SafeOffset = Ptr + Offset
            End If
        End If
    End If

End Function

