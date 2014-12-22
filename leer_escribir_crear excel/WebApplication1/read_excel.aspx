<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="read_excel.aspx.vb" Inherits="WebApplication1.read_excel" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
                             <script type="text/javascript">
                                function alertMe(val) {
                                if (val == 1)   
                                {
                                alert('se proceso correctamente los datos');
                                }
                                 else  if(val==3){
                                alert('no hay registros');
                                }
                                else
                                {
                                alert('error :');
                                }
                                    
                                }
                            </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:FileUpload ID="FileUpload1" runat="server" />
    
    </div>
    <asp:Button ID="btn_ejecutar" runat="server" Text="ejecutar" />
    <p>
        <asp:Label ID="Lbl_resp" runat="server" Text="**"></asp:Label>
    </p>
    <p>
        <asp:Button ID="btn_bajar" runat="server" Text="Descargar Excel" 
            Visible="False" />
    </p>
    </form>
</body>
</html>
