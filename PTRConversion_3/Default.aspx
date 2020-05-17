<%@ Page Title="PTR" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="PTRConversion_3._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <script>
        function pressed(fileInput) {
            var files = fileInput.files;
            var filenameText = document.getElementById('spanPrompt');
            filenameText.innerHTML = files[0].name;
            console.log(filenameText.innerHTML);
        }
    </script>
    <div class="container">
        <section class="section">
            <div class="tile is-ancestor">
                <div class="tile is-4 is-vertical is-parent">
                    <div class="tile is-child box" id="divUpload">
                        <p class="title">Upload</p>
                        <div class="columns is-multiline is-mobile buttons">
                            <div class="column">
                                <div class="file has-name is-boxed" id="inputBox" runat="server">
                                    <label class="file-label" style="min-width: 15rem; max-width: 15rem">
                                        <input class="file-input" type="file" name="FileUploadControl" id="FileUploadControl" runat="server" enabled="true" onchange="pressed(this)" accept=".xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></input>
                                        <span class="file-cta">
                                            <span class="file-icon">
                                                <i class="fas fa-upload"></i>
                                            </span>
                                            <span class="file-label">Browse…
                                            </span>
                                        </span>
                                        <span class="file-name" id="spanPrompt">Please upload a file.</span>
                                    </label>
                                </div>
                                <%--<asp:FileUpload ID="FileUploadControl" runat="server" Enabled="true" />--%>
                            </div>
                            <div class="column is-half">
                                <div class="control">
                                    <a class="button is-link is-medium" id="btnUpload" runat="server" onserverclick="BtnUpload_Click">Upload</a>
                                </div>
                                <%--<asp:Button class="button is-link is-medium" runat="server" ID="btnUpload" Text="Upload" OnClick="BtnUpload_Click" />--%>
                            </div>
                        </div>
                    </div>
                    <div class="tile is-child box">
                        <p class="title">Convert</p>
                        <div class="columns is-multiline is-mobile buttons">
                            <div class="column is-half">
                                <a class="button is-primary is-medium" runat="server" id="btnConvert" text="Convert!" onserverclick="BtnConvert_Click" disabled>Convert!</a>
                            </div>
                            <div class="column is-half">
                                <a class="button is-text is-medium" runat="server" id="btnReset" text="Reset" onserverclick="BtnReset_Click">Reset</a>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="tile is-8 is-parent">
                    <div class="tile is-child box">
                        <p class="title">Output</p>
                        <div class="columns is-vcentered">
                            <div class="column is-full">
                                <article class="message is-link is-large" id="welcomeMessage" runat="server">
                                    <div class="message-body">
                                        Please upload a valid xlsx file through the upload panel.
                                    </div>
                                </article>
                                <article class="message is-success is-large" id="statusMessage" runat="server" visible="false">
                                    <div class="message-body">
                                        <asp:Label runat="server" ID="StatusLabel" Text="Upload status: " />
                                    </div>
                                </article>
                                <article class="message is-danger is-large" id="errorMessage" runat="server" visible="false">
                                    <div class="message-body">
                                        <asp:Label runat="server" ID="ErrorLabel" Text="Error!" />
                                    </div>
                                </article>
                                <div class="buttons is-right">
                                    <a class="button is-warning is-large" runat="server" id="btnDownload" text="Download" onserverclick="BtnDownload_Click" disabled>Download</a>
                                </div>
                                <%--<a class="button is-warning is-large" runat="server" id="btnDownload" text="Download" onserverclick="BtnDownload_Click" disabled>Download</a>--%>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </section>
    </div>
</asp:Content>
