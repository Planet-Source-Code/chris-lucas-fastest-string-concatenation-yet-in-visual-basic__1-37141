VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFastCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ==========================================================================================
' FastCat Class
' Written by Chris Lucas
' E-mail: cdl1051@earthlink.net
'
' For non-commercial usage only! Contact the author for information about commercial
' licenses.
' ==========================================================================================

' Default buffer characteristics, feel free to change these values to better
' suit your needs
Const m_def_BUFFER_PAGE_SIZE As Long = 4096     ' Default page size in bytes
Const m_def_BUFFER_PAGES As Long = 99           ' Default number of buffer pages

' Object level member variables
Private m_BUFFER_PAGE_SIZE As Long              ' Current page size in bytes
Private m_BUFFER_PAGES                          ' Current number of buffer pages

' Object level variables
Private strPages() As String                    ' Dynamic array of buffer pages
Private CreationCount As Long                   ' Number of buffer pages created
Private PageCount As Long                       ' Number of buffer pages used
Private pos As Long                             ' Position in current buffer page




' ==========================================================================================
' ================================= Object Initialization ==================================
' ==========================================================================================
Private Sub Class_Initialize()
    ' Dimension the array of pages
    ReDim strPages(0 To m_def_BUFFER_PAGES)
    
    ' Allocate the first buffer page
    strPages(PageCount) = Space$(m_def_BUFFER_PAGE_SIZE \ 2)
    CreationCount = CreationCount + 1
End Sub



' ==========================================================================================
' ================================= Read / Write Properties ================================
' ==========================================================================================
Friend Property Let BufferPages(NumberOfPages As Long)
    If NumberOfPages > 0 Then
        m_BUFFER_PAGES = NumberOfPages
        ReDim Preserve strPages(0 To m_BUFFER_PAGES - 1)
    Else
        Err.Raise 7776, "cFastCat", "Buffer must contain at least one page"
    End If
End Property

Friend Property Get BufferPages() As Long
Attribute BufferPages.VB_Description = "Returns/sets the number of pages the buffer is allowed to allocate."
    BufferPages = m_BUFFER_PAGES
End Property

Friend Property Let BufferPageSize(PageSize As Long)
    If PageSize > 0 Then
        m_BUFFER_PAGE_SIZE = PageSize
        strPages(PageCount) = Space$(m_BUFFER_PAGE_SIZE \ 2)
        CreationCount = 1
    Else
        Err.Raise 7777, "cFastCat", "Buffer page size must be larger than zero characters"
    End If
End Property

Friend Property Get BufferPageSize() As Long
Attribute BufferPageSize.VB_Description = "Returns/sets the size in bytes of each buffer page."
    BufferPageSize = m_BUFFER_PAGE_SIZE
End Property



' ==========================================================================================
' ================================== Read Only Properties ==================================
' ==========================================================================================
Public Property Get Value() As String
Attribute Value.VB_Description = "Returns the concatenated string"
Attribute Value.VB_UserMemId = 0
    Dim i As Long
    
    Value = Space$((PageCount * m_BUFFER_PAGE_SIZE + pos) \ 2)
    If LenB(Value) = 0 Then Exit Property
    For i = 0 To PageCount
        MidB$(Value, i * m_BUFFER_PAGE_SIZE + 1, m_BUFFER_PAGE_SIZE) = strPages(i)
    Next i
End Property

' Returns the total memory consumed in bytes
Public Property Get TotalBufferSize() As Long
Attribute TotalBufferSize.VB_Description = "Returns the total memory in bytes that the buffer is currently using."
    TotalBufferSize = CreationCount * m_BUFFER_PAGE_SIZE
End Property

Public Property Get StringLength() As Long
Attribute StringLength.VB_Description = "Returns the length of the string contained in the buffer."
    StringLength = PageCount * m_BUFFER_PAGE_SIZE \ 2 + pos
End Property

Public Property Get StringLengthB() As Long
Attribute StringLengthB.VB_Description = "Returns the length of the string in bytes currently contained in the buffer."
    StringLengthB = (PageCount * m_BUFFER_PAGE_SIZE \ 2 + pos) * 2
End Property



' ==========================================================================================
' ================================== Concatenation Methods =================================
' ==========================================================================================
' This sub makes use of a paged buffer and create-on-demand buffer pages.  Both of which
' radically improve string concatenation speed regardless of length
Friend Sub Append(strAppend As String)
Attribute Append.VB_Description = "Concatenates StringData to the end of the string contained in the buffer."
    Dim lng As Long
    Dim partial As Long
    
    lng = LenB(strAppend)

    If pos + lng <= m_BUFFER_PAGE_SIZE Then
        ' Room to write
        MidB$(strPages(PageCount), pos + 1, lng) = strAppend
        pos = pos + lng
    Else
        ' Not enough room, fill up the current buffer page
        partial = m_BUFFER_PAGE_SIZE - pos
        If partial > 0 Then MidB$(strPages(PageCount), pos + 1, partial) = LeftB$(strAppend, partial)
        
        ' Allocate a new buffer page only if one has not previously been allocated
        ' looking before allocating saves significant clock on long concatenations
        PageCount = PageCount + 1
        
        ' Allocation only occurs if there are no more created and unused buffer pages
        If PageCount >= CreationCount Then
            strPages(PageCount) = Space$(m_BUFFER_PAGE_SIZE \ 2)
            CreationCount = CreationCount + 1
        End If
        pos = 0
        
        ' Write the remainder to the new buffer page
        partial = lng - partial
        MidB$(strPages(PageCount), pos + 1, partial) = RightB$(strAppend, partial)
        pos = pos + partial
        
    End If
End Sub

Friend Function ReleaseMemory()
Attribute ReleaseMemory.VB_Description = "Empties the value of the string in the buffer and releases all memory currently used."
    ReDim strPages(0)
    CreationCount = 1
    PageCount = 0
    pos = 0
End Function

Friend Sub Flush()
Attribute Flush.VB_Description = "Empties the value of the string in the buffer.  This method does not release previously allocated buffer pages."
    PageCount = 0
    pos = 0
End Sub
