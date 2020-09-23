<div align="center">

## Return Only Page Name


</div>

### Description

It just gets the page name of your page without any slashes or .asp commented for beginners, 4 lines of code. This is a short little snippet I used today on a project, and didn't see anything like it on PSC, so just thought I would share it.

Comments welcome.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[snowboardr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/snowboardr.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/snowboardr-return-only-page-name__4-8540/archive/master.zip)





### Source Code

```
<%
ScriptName = "/asdfaf/asdf/sfsdf/fas/The_Page_Name.asp"
Function GetPageName(sScriptName)
	'# Location of last slash In ScriptName + 1 because we don't want the slash
	LastSlashLocation = InstrRev(sScriptName,"/") + 1
	'# Location of last period in ScriptName
	LastDotLocation = InstrRev(sScriptName,".")
	'# Here we find out how long the page name is
	ScriptPageNameLength = LastDotLocation - LastSlashLocation
	'# Start 1 after the forward slash, and end before the period.
	'# Set it to the function by using the function name
	GetPageName = Mid(sScriptName,LastSlashLocation,ScriptPageNameLength)
End Function
'# Call Function
Response.Write(GetPageName(ScriptName))
'# Outputs: The_Page_Name
%>
```

