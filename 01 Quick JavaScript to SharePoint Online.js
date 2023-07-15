var SDR_JS_Version = '2023-07-01 0833';

// idea from: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/upload-a-file-by-using-the-rest-api-and-jquery

function get_File_JavaScript(url){
    let scriptTag = document.createElement("script");
    scriptTag.src = url;
    
    scriptTag.onload = function () {
        console.log('onLoad url: ' + url);
    };
    
    let headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    headTag.appendChild(scriptTag);
}

function get_File_CSS(url){
    let link = document.createElement("link");
    link.href = url;
    link.type = "text/css";
    link.rel = "stylesheet";
    link.media = "screen,print";

    let headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    headTag.appendChild(link);
}

// Get the local file as an array buffer.
function getFileBuffer(p_fileInput) {
    var deferred = jQuery.Deferred();

    var reader = new FileReader();

    reader.onloadend = function (e) {
        deferred.resolve(e.target.result);
    }

    reader.onerror = function (e) {
        deferred.reject(e.target.error);
    }

    reader.readAsArrayBuffer(p_fileInput[0].files[0]);

    return deferred.promise();
}

// p_webAbsoluteUrl = _spPageContextInfo.webAbsoluteUrl;
function get_FormDigestValue(p_webAbsoluteUrl){
	let result = '';
	
    let obj = {};
    
    $.ajax({
        url: p_webAbsoluteUrl + '/_api/contextinfo',
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        },
        async:false,
        beforeSend: function() {
            console.log(new Date().toISOString() + ' Start: REST api ContextInfo');
        },
        complete: function() {
            console.log(new Date().toISOString() + ' Complete: REST api ContextInfo');
        },
        success: function (data) {
            obj = data;
        },
        error: function (error) {
            //obj = error;
        }
    });
    
    result = obj.d.GetContextWebInformation['FormDigestValue'];
	
	return result;
}

// Add the file to the file collection in the Shared Documents folder.
function addFileToFolder(arrayBuffer, p_fileInput, p_webAbsoluteUrl, p_serverRelativeUrlToFolder) {
    let parts = p_fileInput[0].value.split('\\');
    let fileName = parts[parts.length - 1];

    // Construct the endpoint.
    var url = `${p_webAbsoluteUrl}/_api/web/getfolderbyserverrelativeurl('${p_serverRelativeUrlToFolder}')/files/add(overwrite=true, url='${fileName}')`;

    return $.ajax({
                url: url,
                type: "POST",
                headers: {
                    "accept": "application/json;odata=verbose"
                    , "X-RequestDigest": get_FormDigestValue(p_webAbsoluteUrl)
                    //, "content-length": arrayBuffer.byteLength
                },
                contentType: "application/json;odata=verbose",
                data: arrayBuffer,
                processData: false,
                beforeSend: function() {
                    console.log(new Date().toISOString() + ' Start: upload file to SharePoint');
                },
                complete: function() {
                    console.log(new Date().toISOString() + ' Complete: upload file to SharePoint');
                }
            });
}

function uploadFile(){
    //let webAbsoluteUrl = _spPageContextInfo.webAbsoluteUrl;
    let webAbsoluteUrl = '';
    let serverRelativeUrlToFolder = 'shared documents';
    let fileInput = jQuery('#getFile');

    if (fileInput[0].files.length <= 0){
        alert('File is empty');

        return;
    }

    getFileBuffer(fileInput).done(function (arrayBuffer){
        let addFile = addFileToFolder(arrayBuffer, fileInput, webAbsoluteUrl, serverRelativeUrlToFolder);
        addFile.done(function (file, status, xhr){
            console.log(new Date().toISOString() + ' Success upload file');
        });
    });
}

/*

var js = 'https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js';

get_File_JavaScript(js);

// clear "article"
$('article').html('');

// create FORM for Upload file
var html_Upload = `
    <div style="margin:20px 50px;">
        <input id="getFile" type="file"/><br /><br />
        <input id="addFileButton" type="button" value="Upload" onclick="uploadFile()"/>
    </div>
`;

$('article').html(html_Upload);

*/
