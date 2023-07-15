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
	
	try
	{
		let obj = {};
		
		$.ajax({
			url: p_webAbsoluteUrl + '/_api/contextinfo',
			type: "POST",
			headers: {
				"accept": "application/json;odata=verbose",
				"content-Type": "application/json;odata=verbose"
			},
            async:false,
			success: function (xyz) {
				obj = xyz;
			},
			error: function (error) {
				//obj = error;
			}
		});
		
		result = obj.d.GetContextWebInformation['FormDigestValue'];
	}
	catch (err){}
	
	return result;
}

// Add the file to the file collection in the Shared Documents folder.
function addFileToFolder(arrayBuffer, p_fileInput, p_webAbsoluteUrl, p_serverRelativeUrlToFolder) {
    // Get the file name from the file input control on the page.
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
                processData: false
            });
}

function uploadFile(){
    //let webAbsoluteUrl = _spPageContextInfo.webAbsoluteUrl;
    let webAbsoluteUrl = '';
    let serverRelativeUrlToFolder = 'shared documents';
    let fileInput = jQuery('#getFile');

    getFileBuffer(fileInput).then(function (arrayBuffer){
        let addFile = addFileToFolder(arrayBuffer, fileInput, webAbsoluteUrl, serverRelativeUrlToFolder);
        addFile.done(function (file, status, xhr){
            console.log('Success upload file');
        });
    });
}

/*

var js = 'https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js';

get_File_JavaScript(js);

// clear "Content"
$('article').html('');

var html_Upload = `
    <div>
        <input id="getFile" type="file"/><br />
        <input id="displayName" type="text" value="Enter a unique name" /><br />
        <input id="addFileButton" type="button" value="Upload" onclick="uploadFile()"/>
    </div>
`;

$('article').html(html_Upload);

*/