<div align="center">

## WraptIt


</div>

### Description

This is a function to provide word wrap capability in your ASP pages.
 
### More Info
 
Function takes a string, a column number, and a forced break boolean. If the forced break boolean is FALSE (default) then the string will wrap at the nearest end-of-word to the column number. If TRUE then the string will wrap at the column number regardless of end-of-word

A string with embedded <br> tags to provide word wraps.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chip Taylor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chip-taylor.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chip-taylor-wraptit__4-6220/archive/master.zip)





### Source Code

```
Function WrapIt(varString,varNumChar,boolBreak)
' Word Wraps a string varString at end of word nearest column position varNumChar
' If boolBreak is true then the wrap will take place at position varNumChar regardless
    varString=Trim(varString)
     varSetBR = "<br>"
    result=""
    varString=Replace(varString,chr(13)+chr(10),chr(10))
    while varString<>""
        if len(varString)<=varNumChar then
            line=varString
            varString=""
        else
            varTemp1=left(varString,varNumChar)
            iCount=InStrRev(varTemp1," ")
            varInPlace=InStr(varTemp1,"chr(10)")
            if (iCount=0) and not boolBreak then iCount=InStr(varString," ")
            if (varInPlace=0) and not boolBreak then varInPlace=InStr(varString,chr(10))
            if (varInPlace<iCount) and (varInPlace>0) then
                finish=j-1
                start=varInPlace+1
            elseif iCount=0 then
                finish=varNumChar
                start=varNumChar+1
            else
                finish=iCount
                start=iCount+1
            end if
            line=left(varString,finish)
            varString=ltrim(mid(varString,start))
        end if
        'if result="" then result=line else
result=result&chr(10)&line
    wend
	   if left(result,1) = chr(10) then
		  result = right(result,(len(result)-1))
		 end if
    WrapIt=Replace(result,chr(10),varSetBR)
end function
```

