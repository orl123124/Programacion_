<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="WebApplication1._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
    
               
      
        
        
        <asp:FileUpload ID="FileUpload1"   runat="server"  />
        
        <asp:Button ID="Button1" runat="server" Text="Ejecutar" />
    
    </div>
    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" 
        ControlToValidate="FileUpload1" ErrorMessage="Archivo no valido, debe elegir un archivo Microsft Excel" 
        ValidationExpression="^.*\.(xlsx|XLSX|XLS|xls)$"></asp:RegularExpressionValidator>
    <p>
        <asp:Button ID="bt_exportar" runat="server" Text="Exportar" />
        <asp:Label ID="lbl_resp" runat="server" Text="*"></asp:Label>
    </p>
    <asp:LinkButton ID="lnk_exp" runat="server">LinkButton</asp:LinkButton>
      
        
       
    </form>
</body>
</html>
