/*
    Author      : SDR - Software DevelopeR
    Description : utility related with SharePoint REST API
*/

// https://github.com/SharePoint/sp-dev-docs/blob/main/docs/spfx/overview-graphhttpclient.md
//      To acquire a valid access token,
//      GraphHttpClient issues a web request to the /_api/SP.OAuth.Token/Acquire endpoint. 
//      This API is intended for internal use. You should not communicate with it directly in your solutions.
//
async function SDR_AccessToken(){
    let deferred = $.Deferred();
    
    let webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let resourceData = {
        "resource": "https://graph.microsoft.com"
        //"resource": "https://outlook.office365.com/search"
        //"resource": _spPageContextInfo.webAbsoluteUrl
    };
    
    let url = webUrl + "/_api/SP.OAuth.Token/Acquire";
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(resourceData)
    })
    .done(function(result) {
        let access_token = result.d.access_token;
        
        deferred.resolve(access_token);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    var data = await SDR_AccessToken();
    
    console.log(data);
})();

*/

function validate_AccessToken(accessToken){
    let deferred = $.Deferred();
    
    let url = "https://graph.microsoft.com/v1.0/me";
    
    $.ajax({
        url: url,
        type: "GET",
        headers: {
            'Accept': 'application/json',
            'Authorization': 'Bearer ' + accessToken
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    var accessToken = await SDR_AccessToken();
    var data = await validate_AccessToken(accessToken);
    
    console.log(data);
})();

*/

async function accessToken_Today(){
    let accessToken = '';
    
    try{
        //let date = new Date().toISOString().replace(/[\-\:\.]/g, "");
        let date = new Date().toISOString();
        //let today = date.substring(0, 10);
        let today = date.substring(0, 13);
        
        let key = "AccessToken_" + today;
        
        accessToken = localStorage.getItem(key);
        if ( ! accessToken){
            accessToken = await SDR_AccessToken();
            
            localStorage.setItem(key, accessToken);
        }
    }catch(e){}
    
    return accessToken;
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    var data = await accessToken_Today();
    
    console.log(data);
})();

*/

function formDigestValueAsync(webUrl){
    let deferred = $.Deferred();
    
    let url = webUrl + "/_api/contextinfo";
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        try{
            let formDigestValue = result.d.GetContextWebInformation['FormDigestValue'];
        
            deferred.resolve(formDigestValue);
        }catch(e){
            deferred.reject('');
        }
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var data = await formDigestValueAsync(webUrl);
    
    console.log(data);
})();

*/

var delayFunc = ms => new Promise(res => setTimeout(res, ms));

// seemingly: this does not work
function sleeperFunc(ms) {
    return function(x) {
        return new Promise(resolve => setTimeout(() => resolve(x), ms));
    };
}

function waitformeFunc(delay_millisec) {
    return new Promise(resolve => {
        setTimeout(() => { resolve('') }, delay_millisec);
    });
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    console.log('-- delayFunc()');
    console.log('Delay 2.000 ms ...');
    await delayFunc(2000);
    console.log('Complete.');
    
    console.log('Delay 3.000 ms ...');
    await delayFunc(3000);
    console.log('Complete.');
    
    console.log('Delay 5.000 ms ...');
    await delayFunc(5000);
    console.log('Complete.');
    
    console.log('-- sleeperFunc()');
    console.log('Delay 2.000 ms ...');
    await sleeperFunc(2000);
    console.log('Complete.');
    
    console.log('Delay 3.000 ms ...');
    await sleeperFunc(3000);
    console.log('Complete.');
    
    console.log('Delay 5.000 ms ...');
    await sleeperFunc(5000);
    console.log('Complete.');
    
    console.log('-- waitformeFunc()');
    console.log('Delay 2.000 ms ...');
    await waitformeFunc(2000);
    console.log('Complete.');
    
    console.log('Delay 3.000 ms ...');
    await waitformeFunc(3000);
    console.log('Complete.');
    
    console.log('Delay 5.000 ms ...');
    await waitformeFunc(5000);
    console.log('Complete.');
    
})();

*/

// https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
// https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints
//
function getListItemsAsync(webUrl, listName, fields, filter, limit, orderby){
    let deferred = $.Deferred();
    
    let param_select = '$select=*';
    if (fields){
        param_select = '$select=' + fields;
    }
    
    let param_top = '';
    if (limit){
        param_top = '&$top=' + limit;
    }
    
    // https://support.shortpoint.com/support/solutions/articles/1000307202-shortpoint-rest-api-selecting-filtering-sorting-results-in-a-sharepoint-list
    
    let param_filter = '';
    if (filter){
        param_filter = '&$filter=' + filter;
    }
    
    let param_orderby = '';
    if (orderby){
        param_orderby = '&$orderby=' + orderby + ' asc';
    }
    
    let url = webUrl + `/_api/web/lists/getbytitle('${listName}')/items?${param_select}${param_top}${param_filter}${param_orderby}`;
    
    $.ajax({
        url : url,
        type: "GET",
        contentType : "application/json;odata=verbose",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

async function get_AllFields(webUrl, listName){
    let result = {};
    
    let fields = '*,FieldValuesAsText&$expand=FieldValuesAsText';
    let filter = '';
    let limit = 1;
    
    let data = await getListItemsAsync(webUrl, listName, fields, filter, limit);
    
    try
    {
        let results = data.d.results;
        
        let FieldValues = results[0]["FieldValuesAsText"];  
        result = Object.keys(FieldValues);
    }catch(e){}
    
    return result;
}

async function get_Type_SPdata_ori(webUrl, listName){
    let result = '';
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')`;
    
    let data = await getData_GET(url);
    
    let entityTypeName = data.d.EntityTypeName;
    
    result = `SP.Data.${entityTypeName}Item`;
    
    return result;
}

async function get_Type_SPdata(webUrl, listName){
    // 1. this is just "Quick Way"
    // 2. let's use function get_Type_SPdata_ori() for accurate value {just in case listName is renamed }
    let result = `SP.Data.${listName}ListItem`;
    
    return result;
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var fields = '';
    //var fields = 'ID,Name,Created';
    //var fields = '*,FieldValuesAsText&$expand=FieldValuesAsText';
    
    var filter = '';
    //var filter = 'ID eq 1';
    //var filter = "startswith(Name, 'test')";
    //var filter = "substringof('test', Name)";
    
    var limit = '';
    var orderby = '';
    //var orderby = 'Name';
    
    var data = await getListItemsAsync(webUrl, listName, fields, filter, limit, orderby);
    
    console.log(data);
})();

(async () => {
    var listName = 'Person';
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var data = await get_AllFields(webUrl, listName);
    
    console.log(data);
})();

(async () => {
    var listName = 'Person';
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var data = await get_Type_SPdata(webUrl, listName);
    
    console.log(data);
})();

*/

function getData_GET(url){
    let deferred = $.Deferred();
    
    $.ajax({
        url: url,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    var listName = 'Person';
    
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/items?$select=ID,Name,Created`;
    
    var data = await getData_GET(url);
    
    console.log(data);
})();

(async () => {
    // kolom AttachmentFiles
    // di 2 tempat, yaitu:
    //      1. $select=AttachmentFiles
    //      2. $expand=AttachmentFiles
    var url = "/sites/Workspaces/Testing/_api/lists/getByTitle('List with Attachment')/items?$select=AttachmentFiles,Title,Age,Skills&$expand=AttachmentFiles";
    
    var data = await getData_GET(url);
    
    console.log(data);
})();

*/

function getData_POST(url, data){
    let deferred = $.Deferred();
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        },
        data: JSON.stringify(data)
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

// https://learn.microsoft.com/en-us/answers/questions/787712/retrieve-all-fields-in-sharepoint-list-using-jquer

function allData_toCSV(data){
    let results = data.d.results;
        
    let FieldValues = results[0]["FieldValuesAsText"];  
    let fields = Object.keys(FieldValues);  
    
    let csv = "data:text/csv;charset=utf-8," + fields.join(",") + "\n";  
    for (let j = 0; j < results.length; j++) {  

        for (let k = 0; k < fields.length; k++) {  
            csv += results[j][fields[k]];  
            csv += k < fields.length - 1 ? "," : "";  
        }  
        csv += "\n";  
    }  
    
    let a = document.createElement("a");  
    a.setAttribute("href", encodeURI(csv));  
    a.setAttribute("download", "DataList-" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".csv");  
    document.body.appendChild(a);  
    a.click();  
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var listName = 'Person';
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var fields = '*,FieldValuesAsText&$expand=FieldValuesAsText';
    
    var data = await getListItemsAsync(webUrl, listName, fields);
    
    console.log(data);
    
    allData_toCSV(data);
})();

*/

async function sendEmailAsync(from, to, body, subject, webUrl){
    let deferred = $.Deferred();
    
    if ( ! webUrl){
        webUrl = _spPageContextInfo.webServerRelativeUrl;
    }
    
    let url = webUrl + "/_api/SP.Utilities.Utility.SendEmail";
    
    let to_Array = [];
    if ( ! Array.isArray(to)){
        to_Array.push(to);
    }else{
        to_Array = to;
    }
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify({
            'properties': {
                '__metadata': {
                    'type': 'SP.Utilities.EmailProperties'
                },
                'From': from,
                'To': {
                    'results': to_Array
                },
                'Body': body,
                'Subject': subject
            }
        })
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var from = 'admin@platform365.onmicrosoft.com';
    var to = 'softwaredeveloper@platform365.onmicrosoft.com';
    var body = 'Test only';
    var subject = 'Test on - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var data = await sendEmailAsync(from, to, body, subject);
    
    console.log(data);
})();

(async () => {
    var from = 'admin@platform365.onmicrosoft.com';
    var to = ['softwaredeveloper@platform365.onmicrosoft.com'
                , 'admin@platform365.onmicrosoft.com'
                , 'test@platform365.onmicrosoft.com'
                , 'user1@platform365.onmicrosoft.com'
                , 'user2@platform365.onmicrosoft.com'];
    var body = 'Test only';
    var subject = 'Test on - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var data = await sendEmailAsync(from, to, body, subject);
    
    console.log(data);
})();

*/

async function addListItemAsync(webUrl, listName, dataObj){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/items`;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    let listItem_Type = await get_Type_SPdata(webUrl, listName);
    
    dataObj['__metadata'] = {'type': listItem_Type};
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(dataObj)
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var languageValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["Indonesia"]};
    var skillValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["JavaScript", "PHP", "C#", "HTML"]};
    var friendValues = {"__metadata":{"type":"Collection(Edm.Int32)"},"results":[18, 19]};
    
    var title = 'Test on - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var dataObj = {"Title": title,
                    "Name": "Test only",
                    "Gender": "Male",                       // Choice, single value
                    "Language": languageValues,             // Choice, multiple value
                    "Skills": skillValues,                  // Choice, multiple value
                    "Birthday": new Date().toISOString(),   // Date and Time
                    "Experience": 123,
                    "HealthInsurance": true,                // Yes/No
                    "WilayahId": 2,                         // Lookup
                    "ManagerId": 15,                        // Person or Group
                    "FriendsId": friendValues               // Person or Group, multiple value
    };
    
    var data = await addListItemAsync(webUrl, listName, dataObj);
    
    console.log(data);
})();

*/

async function updateListItemAsync(webUrl, listName, id, dataObj){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/items(${id})`;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    let listItem_Type = await get_Type_SPdata(webUrl, listName);
    
    dataObj['__metadata'] = {'type': listItem_Type};
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(dataObj)
    })
    .done(function(result) {
        if (result){
            deferred.resolve(result);
        }else{
            deferred.resolve(true);
        }
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var languageValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["English"]};
    var skillValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["Reading", "Writing"]};
    var friendValues = {"__metadata":{"type":"Collection(Edm.Int32)"},"results":[18, 19]};
    
    var title = 'Test on - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var dataObj = {"Title": title,
                    "Name": "Test only - changed",
                    "Gender": "Female",                     // Choice, single value
                    "Language": languageValues,             // Choice, multiple value
                    "Skills": skillValues,                  // Choice, multiple value
                    "Birthday": new Date().toISOString(),   // Date and Time
                    "Experience": 23,
                    "HealthInsurance": true,                // Yes/No
                    "WilayahId": 2,                         // Lookup
                    "ManagerId": 15,                        // Person or Group
                    "FriendsId": friendValues               // Person or Group, multiple value
    };
    
    var data = await updateListItemAsync(webUrl, listName, 60, dataObj);
    
    console.log(data);
})();

*/

async function deleteListItemAsync_simple(webUrl, listName, id){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/items(${id})`;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        }
    })
    .done(function(result) {
        if (result){
            deferred.resolve(result);
        }else{
            deferred.resolve(true);
        }
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var data = await deleteListItemAsync_simple(webUrl, listName, 1);
    
    console.log(data);
})();

*/

async function deleteListItemAsync_Digest(webUrl, listName, id, formDigestValue){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/items(${id})`;
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        }
    })
    .done(function(result) {
        if (result){
            deferred.resolve(result);
        }else{
            deferred.resolve(true);
        }
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

async function deleteListItemAsync_Digest_X(webUrl, listName, id, formDigestValue){
    let result = false;
    
    try{
        let data = await deleteListItemAsync_Digest(webUrl, listName, id, formDigestValue);
        
        result = true;
    }catch(e){}
    
    return result;
}

async function deleteListItemAsync(webUrl, listName, id){
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let data = await deleteListItemAsync_Digest(webUrl, listName, id, formDigestValue);
    
    return data;
}

async function deleteListItemAsync_X(webUrl, listName, id){
    let result = false;
    
    try{
        let data = await deleteListItemAsync(webUrl, listName, id);
        
        result = true;
    }catch(e){}
    
    return result;
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var data = await deleteListItemAsync(webUrl, listName, 1);
    
    console.log(data);
})();

*/

async function deleteListItemsAsync_N(webUrl, listName, id_Array){
    let to_Array = [];
    if ( ! Array.isArray(id_Array)){
        to_Array.push(id_Array);
    }else{
        to_Array = id_Array;
    }
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let data = {};
    
    for(let i=0;i<to_Array.length;i++){
        data = await deleteListItemAsync_Digest_X(webUrl, listName, to_Array[i], formDigestValue);
    }
    
    return true;
}

function arrayRange(start, stop, step){
    return Array.from(
                        { length: (stop - start) / step + 1 },
                            (value, index) => start + index * step
    );
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var ids = arrayRange(1, 100, 1);
    var data = await deleteListItemsAsync_N(webUrl, listName, ids);
    
    console.log(data);
})();

*/

// https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#create-list-item-in-a-folder
async function addFolder_toList(webUrl, listName, folderName){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/AddValidateUpdateItemUsingPath`;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let dataObj = {
                    "listItemCreateInfo": {
                        "FolderPath": {
                            "DecodedUrl": webUrl + `/Lists/${listName}`
                        },
                        "UnderlyingObjectType": 1 // FileSystemObjectType {Folder = 1, File = 0}
                    },
                    "formValues": [{
                                    "FieldName": "Title",
                                    "FieldValue": folderName
                                    }],
                    "bNewDocumentUpdate": false
    };
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(dataObj)
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'List with Folder';
    var folderName = 'Folder - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var data = await addFolder_toList(webUrl, listName, folderName);
    
    console.log(data);
})();

*/

async function addItem_toFolder_inList(webUrl, listName, folderName, dataObj_Item){
    // first Step:
    //      - add to certain 'Folder' in List
    //      - but just column: 'Title' only
    //      - then change values of other columns
    let newItemId = await addItem_toFolder_inList_firstStep_Title_only(webUrl, listName, folderName, dataObj_Item);
    
    // second Step:
    //      - change values of other columns
    let data = await updateListItemAsync(webUrl, listName, newItemId, dataObj_Item);
    
    return data;
}

async function addItem_toFolder_inList_firstStep_Title_only(webUrl, listName, folderName, dataObj_Item){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/AddValidateUpdateItemUsingPath`;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let dataObj = {
                    "listItemCreateInfo": {
                        "FolderPath": {
                            "DecodedUrl": webUrl + `/Lists/${listName}/${folderName}`
                        },
                        "UnderlyingObjectType": 0 // FileSystemObjectType {Folder = 1, File = 0}
                    },
                    "formValues": [{
                                    "FieldName": "Title",
                                    "FieldValue": dataObj_Item.Title
                                    }],
                    "bNewDocumentUpdate": false
    };
    
    let newObject = {};
    let newItemId = 0;
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(dataObj)
    })
    .done(function(result) {
        newObject = result;
        
        let newItems = newObject.d.AddValidateUpdateItemUsingPath.results;
    
        for (let i=0;i<newItems.length;i++){
            if (newItems[i].FieldName == "Id"){
                newItemId = newItems[i].FieldValue;
                break;
            }
        }
        
        deferred.resolve(newItemId);
        
        //deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var languageValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["Indonesia"]};
    var skillValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["JavaScript", "PHP", "C#", "HTML"]};
    var friendValues = {"__metadata":{"type":"Collection(Edm.Int32)"},"results":[18, 19]};
    
    var title = 'Test on - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var dataObj = {"Title": title,
                    "Name": "Test only",
                    "Gender": "Male",                       // Choice, single value
                    "Language": languageValues,             // Choice, multiple value
                    "Skills": skillValues,                  // Choice, multiple value
                    "Birthday": new Date().toISOString(),   // Date and Time
                    "Experience": 123,
                    "HealthInsurance": true,                // Yes/No
                    "WilayahId": 2,                         // Lookup
                    "ManagerId": 15,                        // Person or Group
                    "FriendsId": friendValues               // Person or Group, multiple value
    };
    
    var folderName = "Test";
    
    var newItem = await addItem_toFolder_inList(webUrl, listName, folderName, dataObj);
    
    console.log(newItem);
})();

*/

// https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#create-list-item-in-a-folder
// I still do not understand, how to call AddValidateUpdateItemUsingPath
//      does not work for some columns:
//              - Person or Group
//              - Yes/No
//      it works well for columns:
//              - Choice, multiple value
//              - Lookup
//
// Note: use function addItem_toFolder_inList() instead
//
async function addItem_toFolder_inList_Not_Work(webUrl, listName, folderName, dataObj_Item){
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/lists/GetByTitle('${listName}')/AddValidateUpdateItemUsingPath`;
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let formValues = [];
    
    for (const [key, value] of Object.entries(dataObj_Item)) {
        formValues.push({"FieldName": key,
                         "FieldValue": value
                        });
    }
    
    let dataObj = {
                    "listItemCreateInfo": {
                        "FolderPath": {
                            "DecodedUrl": webUrl + `/Lists/${listName}/${folderName}`
                        },
                        "UnderlyingObjectType": 0 // FileSystemObjectType {Folder = 1, File = 0}
                    },
                    "formValues": formValues,
                    "bNewDocumentUpdate": false
    };
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(dataObj)
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    var listName = 'Person';
    
    var dateValue = new Date().toISOString();
    
    var dateValue = new Date().toLocaleDateString("en-US");
    
    var manager = "[{'Key':'i:0#.f|softwaredeveloper@platform365.onmicrosoft.com'}]";
    var manager = "i:0#.f|softwaredeveloper@platform365.onmicrosoft.com";
    var manager = "softwaredeveloper@platform365.onmicrosoft.com";
    
    var languageValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["Indonesia"]};
    var skillValues = {"__metadata":{"type":"Collection(Edm.String)"},"results":["JavaScript", "PHP", "C#", "HTML"]};
    var friendValues = {"__metadata":{"type":"Collection(Edm.Int32)"},"results":[18, 19]};
    
    // use Simple ;#
    var languageValues = "Indonesia";
    var skillValues = "JavaScript;#PHP;#C#;#HTML";
    var friendValues = "18;#19";
    
    var title = 'Test on - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var dataObj = {"Title": title,
                    "Name": "Test only",
                    "Gender": "Male",                       // Choice, single value
                    "Language": languageValues,             // Choice, multiple value
                    "Skills": skillValues,                  // Choice, multiple value
                    "Birthday": dateValue,
                    "Experience": (123).toString(),
                    "HealthInsurance": "false",             // Yes/No
                    "Wilayah": (2).toString(),              // Lookup
                    "Manager": manager,                     // Person or Group
                    "Friends": friendValues                 // Person or Group, multiple value
    };
    
    var folderName = "Test";
    
    var data = await addItem_toFolder_inList_Not_Work(webUrl, listName, folderName, dataObj);
    
    console.log(data);
})();

*/

function SiteUsers(fields, filter, limit, orderby, webUrl){
    let deferred = $.Deferred();
    
    let param_select = '$select=*';
    if (fields){
        param_select = '$select=' + fields;
    }
    
    let param_top = '';
    if (limit){
        param_top = '&$top=' + limit;
    }
    
    // https://support.shortpoint.com/support/solutions/articles/1000307202-shortpoint-rest-api-selecting-filtering-sorting-results-in-a-sharepoint-list
    
    let param_filter = '';
    if (filter){
        param_filter = '&$filter=' + filter;
    }
    
    let param_orderby = '';
    if (orderby){
        param_orderby = '&$orderby=' + orderby + ' asc';
    }
    
    if ( ! webUrl){
        webUrl = _spPageContextInfo.webServerRelativeUrl;
    }
    
    let url = webUrl + `/_api/web/siteusers?${param_select}${param_top}${param_filter}${param_orderby}`;
    
    $.ajax({
        url: url,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    
    var data = await SiteUsers();
    
    console.log(data);
})();

(async () => {
    
    var data = await SiteUsers('Id, Email, UserPrincipleName, Title', 'Id eq 17 or Id eq 12');
    
    console.log(data);
})();

*/

function SiteUsers_by_Ids(fields, ids, limit, orderby, webUrl){
    let deferred = $.Deferred();
    
    if ( ! ids){
        let result = {
                        "d": {
                                "results": []
                        }
                    };
        
        deferred.resolve(result);
        
        return deferred.promise();
    }
    
    let param_select = '$select=*';
    if (fields){
        param_select = '$select=' + fields;
    }
    
    let param_top = '';
    if (limit){
        param_top = '&$top=' + limit;
    }
    
    let ids_Array = [];
    if ( ! Array.isArray(ids)){
        ids_Array.push(ids);
    }else{
        ids_Array = ids;
    }
    
    let filter = '';
    let separator_Or = '';
    ids_Array.forEach((item) => {
        filter += separator_Or + 'Id eq ' + item.toString();
        separator_Or = ' or ';
    });
    
    let param_filter = '';
    if (filter){
        param_filter = '&$filter=' + filter;
    }
    
    let param_orderby = '';
    if (orderby){
        param_orderby = '&$orderby=' + orderby + ' asc';
    }
    
    if ( ! webUrl){
        webUrl = _spPageContextInfo.webServerRelativeUrl;
    }
    
    let url = webUrl + `/_api/web/siteusers?${param_select}${param_top}${param_filter}${param_orderby}`;
    
    $.ajax({
        url: url,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)
(async () => {
    
    var data = await SiteUsers_by_Ids();
    
    console.log(data);
})();

(async () => {
    
    var data = await SiteUsers_by_Ids('Id, Email, UserPrincipleName, Title', 12);
    
    console.log(data);
})();

(async () => {
    
    var data = await SiteUsers_by_Ids('Id, Email, UserPrincipleName, Title', [12]);
    
    console.log(data);
})();

(async () => {
    
    var data = await SiteUsers_by_Ids('Id, Email, UserPrincipleName, Title', [12, 11, 16]);
    
    console.log(data);
})();

*/

// https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest

async function Folder_Create(webUrl, documentLibrary, folderName){
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let folders = folderName.split('/');
    let n = folders.length;
    
    let createdStatus = false;
    let folder_str = '';
    let separator = '';
    
    // handle Folder Hierarchy
    // example 'abc/def/ghi'
    for (let i=0;i<n;i++){
        folder_str += separator + folders[i];
        
        createdStatus = await Folder_Create_Digest(webUrl, documentLibrary, folder_str, formDigestValue);
        
        separator = '/';
    }
    
    return createdStatus;
}

async function Folder_Create_Digest(webUrl, documentLibrary, folderName, formDigestValue){
    let retFunction = false;
    
    let deferred = $.Deferred();
    
    let url = webUrl + `/_api/web/folders`;
    
    let dataObj = {
                        "__metadata": {
                            "type": "SP.Folder"
                        },
                        // do not forget to add webUrl in parameter ServerRelativeUrl
                        "ServerRelativeUrl": webUrl + `/${documentLibrary}/${folderName}`
                    };
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(dataObj)
    })
    .done(function(result) {
        try{
            if (result.d.Name){
                retFunction = true;
            }
        }catch(e){}
        
        deferred.resolve(retFunction);
    })
    .fail(function(result, status) {
        //deferred.reject(status);
        
        // avoid: "Uncaught (in promise)"
        // resolve with return function: false
        deferred.resolve(retFunction);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Shared Documents';
    var folderName = 'Folder - ' + (new Date().toISOString().replace(/[\-\:\.]/g, ""));
    
    var data = await Folder_Create(webUrl, documentLibrary, folderName);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Shared Documents';
    var folderName = 'abc';
    
    var data = await Folder_Create(webUrl, documentLibrary, folderName);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Shared Documents';
    var folderName = 'abc/def';
    
    var data = await Folder_Create(webUrl, documentLibrary, folderName);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Shared Documents';
    var folderName = 'abc/def/ghi';
    
    var data = await Folder_Create(webUrl, documentLibrary, folderName);
    
    console.log(data);
})();

*/

function Folder_Exists(webUrl, folder, documentLibrary){
    let retFunction = false;
    
    let deferred = $.Deferred();
    
    let vFolder = folder;
    if (documentLibrary){
        vFolder = `${documentLibrary}/${folder}`;
    }
    
    // do not forget to add webUrl in parameter function GetFolderByServerRelativeUrl()
    let url = webUrl + `/_api/web/GetFolderByServerRelativeUrl('${webUrl}/${vFolder}')/Exists`;
    
    $.ajax({
        url: url,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        try{
            retFunction = result.d.Exists;
        }catch(e){}
        
        deferred.resolve(retFunction);
    })
    .fail(function(result, status) {
        //deferred.reject(status);
        
        // avoid: "Uncaught (in promise)"
        // resolve with return function: false
        deferred.resolve(retFunction);
    });
    
    return deferred.promise();
}

/*
// Test: 

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Shared Documents';
    var folderName = 'abc';
    
    var data = await Folder_Exists(webUrl, folderName, documentLibrary);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var folderName = 'Shared Documents/abc';
    
    var data = await Folder_Exists(webUrl, folderName);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var folderName = 'Shared Documents/abc/def/ghi';
    
    var data = await Folder_Exists(webUrl, folderName);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var folderName = 'Shared Documents/x';
    
    var data = await Folder_Exists(webUrl, folderName);
    
    console.log(data);
})();

*/

function getDocumentLibrary_ItemsAsync(webUrl, documentLibrary, fields, filter, limit, orderby){
    let deferred = $.Deferred();
    
    let param_select = '$select=*';
    if (fields){
        param_select = '$select=' + fields;
    }
    
    let param_top = '';
    if (limit){
        param_top = '&$top=' + limit;
    }
    
    let param_filter = '';
    if (filter){
        param_filter = '&$filter=' + filter;
    }
    
    let param_orderby = '';
    if (orderby){
        param_orderby = '&$orderby=' + orderby + ' asc';
    }
    
    // do not forget to add webUrl in parameter function GetFolderByServerRelativeUrl()
    let url = webUrl + `/_api/web/GetFolderByServerRelativeUrl('${webUrl}/${documentLibrary}')?${param_select}${param_top}${param_filter}${param_orderby}`;
    
    $.ajax({
        url : url,
        type: "GET",
        contentType : "application/json;odata=verbose",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose"
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

async function getDocumentLibrary_FilesFoldersAsync(webUrl, documentLibrary, filter, limit, orderby){
    let fields = 'Folders,Files&$expand=Folders,Files';
    
    let data = await getDocumentLibrary_ItemsAsync(webUrl, documentLibrary, fields, filter, limit, orderby);
    
    return data;
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    // DocumentLibrary as List
    var documentLibrary = 'Documents';
    
    var fields = '*,FieldValuesAsText&$expand=FieldValuesAsText';
    
    var data = await getListItemsAsync(webUrl, documentLibrary, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    // DocumentLibrary as List
    var documentLibrary = 'Documents';
    
    var fields = 'FileLeafRef,FileRef,Id';
    
    var data = await getListItemsAsync(webUrl, documentLibrary, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    // DocumentLibrary as DocumentLibrary
    var documentLibrary = 'Shared Documents';
        
    var data = await getDocumentLibrary_ItemsAsync(webUrl, documentLibrary);
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    // DocumentLibrary as DocumentLibrary
    var documentLibrary = 'Shared Documents';
    
    // Files and Folders
    var data = await getDocumentLibrary_FilesFoldersAsync(webUrl, documentLibrary);
    
    console.log(data);
})();

*/

// https://sharepoint.stackexchange.com/questions/138135/get-all-files-and-folders-in-one-call

async function getDocumentLibrary_RecursiveAsync(webUrl, documentLibrary, folder, viewXml, fields, filter, limit, orderby){
    let deferred = $.Deferred();
    
    let param_select = '$select=*';
    if (fields){
        param_select = '$select=' + fields;
    }
    
    let param_top = '';
    if (limit){
        param_top = '&$top=' + limit;
    }
    
    let param_filter = '';
    if (filter){
        param_filter = '&$filter=' + filter;
    }
    
    let param_orderby = '';
    if (orderby){
        param_orderby = '&$orderby=' + orderby + ' asc';
    }
    
    let param_folder = '';
    if (folder){
        param_folder = '/' + folder;
    }
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    // beware:
    //    if DocumentLibrary has 2 different Name
    //          - Title
    //          - Internal Name
    // example:
    //          - Documents
    //          - Shared Documents
    // that different will confuse value in function BeginsWith
    
    // https://sharepoint.stackexchange.com/questions/208020/make-caml-query-with-in-rest-api-call
    
    // https://learn.microsoft.com/en-us/sharepoint/dev/schema/beginswith-element-query
    let v_viewXml = {"ViewXml":
                        `<View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <BeginsWith>
                                        <FieldRef Name='FileDirRef' />
                                        <Value Type='Text'>${webUrl}/${documentLibrary}${param_folder}</Value>
                                    </BeginsWith>
                                </Where>
                            </Query>
                        </View>`};
                        
    if (viewXml){
        v_viewXml = viewXml;
    }
    
    let url = webUrl + `/_api/web/Lists/GetByTitle('${documentLibrary}')/GetItems(query=@v1)?${param_select}${param_top}${param_filter}${param_orderby}&@v1=${JSON.stringify(v_viewXml)}`;
    
    $.ajax({
        url : url,
        type: "POST",
        contentType : "application/json;odata=verbose",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        }
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = '';
    var viewXml = null;
    var fields = 'ID,FileLeafRef,FileDirRef';
    
    var data = await getDocumentLibrary_RecursiveAsync(webUrl, documentLibrary, folder, viewXml, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = 'abc/def/ghi';
    var viewXml = null;
    var fields = 'ID,FileLeafRef,FileDirRef';
    
    var data = await getDocumentLibrary_RecursiveAsync(webUrl, documentLibrary, folder, viewXml, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = 'abc/def/ghi';
    var viewXml = null;
    var fields = 'ID,FileLeafRef,FileDirRef';
    
    var viewXml = {"ViewXml":
                        `<View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <BeginsWith>
                                        <FieldRef Name='FileDirRef' />
                                        <Value Type='Text'>${webUrl}/${documentLibrary}/${folder}</Value>
                                    </BeginsWith>
                                </Where>
                            </Query>
                        </View>`};
    
    var data = await getDocumentLibrary_RecursiveAsync(webUrl, documentLibrary, folder, viewXml, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = 'abc/def/ghi/x1';
    var viewXml = null;
    var fields = 'ID,FileLeafRef,FileDirRef';
    
    var folder2 = 'abc/def/ghi/x3';
    
    var viewXml = {"ViewXml":
                        `<View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <Or>
                                        <BeginsWith>
                                            <FieldRef Name='FileDirRef' />
                                            <Value Type='Text'>${webUrl}/${documentLibrary}/${folder}</Value>
                                        </BeginsWith>
                                        <BeginsWith>
                                            <FieldRef Name='FileDirRef' />
                                            <Value Type='Text'>${webUrl}/${documentLibrary}/${folder2}</Value>
                                        </BeginsWith>
                                    </Or>
                                </Where>
                            </Query>
                        </View>`};
                        
    // Note:
    //      - limitation query in URL
    //      - viewXml will generate Error 404 {because string is too long}
    //      - limitation of REST URL (endpoint) of 260 characters
    // Action:
    //      - rewrite variable as following:
    
    var viewXml = {"ViewXml": `<View Scope='RecursiveAll'><Query><Where><Or><BeginsWith><FieldRef Name='FileDirRef' /><Value Type='Text'>${webUrl}/${documentLibrary}/${folder}</Value></BeginsWith><BeginsWith><FieldRef Name='FileDirRef' /><Value Type='Text'>${webUrl}/${documentLibrary}/${folder2}</Value></BeginsWith></Or></Where></Query></View>`};
                        
    var data = await getDocumentLibrary_RecursiveAsync(webUrl, documentLibrary, folder, viewXml, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    
    console.log(data);
})();

*/

// https://sharepoint.stackexchange.com/questions/208020/make-caml-query-with-in-rest-api-call
//  - limitation of REST URL (endpoint) of 260 characters
//  - Caml Query in Body
async function getDocumentLibrary_RecursiveAsync_inBody(webUrl, documentLibrary, folder, viewXml, fields, filter, limit, orderby){
    let deferred = $.Deferred();
    
    let param_select = '$select=*';
    if (fields){
        param_select = '$select=' + fields;
    }
    
    let param_top = '';
    if (limit){
        param_top = '&$top=' + limit;
    }
    
    let param_filter = '';
    if (filter){
        param_filter = '&$filter=' + filter;
    }
    
    let param_orderby = '';
    if (orderby){
        param_orderby = '&$orderby=' + orderby + ' asc';
    }
    
    let param_folder = '';
    if (folder){
        param_folder = '/' + folder;
    }
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    // beware:
    //    if DocumentLibrary has 2 different Name
    //          - Title
    //          - Internal Name
    // example:
    //          - Documents
    //          - Shared Documents
    // that different will confuse value in function BeginsWith
    
    let v_viewXml = {"ViewXml":
                        `<View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <BeginsWith>
                                        <FieldRef Name='FileDirRef' />
                                        <Value Type='Text'>${webUrl}/${documentLibrary}${param_folder}</Value>
                                    </BeginsWith>
                                </Where>
                            </Query>
                        </View>`};
                        
    if (viewXml){
        v_viewXml = viewXml;
    }
    
    
    var data = { 'query':
                            {
                                '__metadata': { 'type': 'SP.CamlQuery' },
                                'ViewXml': v_viewXml.ViewXml
                            }
                };
                        
    
    let url = webUrl + `/_api/web/Lists/GetByTitle('${documentLibrary}')/GetItems?${param_select}${param_top}${param_filter}${param_orderby}`;
    
    $.ajax({
        url : url,
        type: "POST",
        contentType : "application/json;odata=verbose",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: JSON.stringify(data)
    })
    .done(function(result) {
        deferred.resolve(result);
    })
    .fail(function(result, status) {
        deferred.reject(status);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = 'abc/def/ghi/x1';
    var viewXml = null;
    var fields = 'ID,FileLeafRef,FileDirRef';
    
    var folderB = 'abc/def/ghi/x3';
    var folderC = 'abc/def/ghi/x5';
    
    var viewXml = {"ViewXml":
                        `<View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <Or>
                                        <BeginsWith>
                                            <FieldRef Name='FileDirRef' />
                                            <Value Type='Text'>${webUrl}/${documentLibrary}/${folder}</Value>
                                        </BeginsWith>
                                        <BeginsWith>
                                            <FieldRef Name='FileDirRef' />
                                            <Value Type='Text'>${webUrl}/${documentLibrary}/${folderB}</Value>
                                        </BeginsWith>
                                    </Or>
                                </Where>
                            </Query>
                        </View>`};
                        
    var data = await getDocumentLibrary_RecursiveAsync_inBody(webUrl, documentLibrary, folder, viewXml, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = 'abc/def/ghi/x1';
    var viewXml = null;
    var fields = 'ID,FileLeafRef,FileDirRef';
    
    var folderB = 'abc/def/ghi/x3';
    var folderC = 'abc/def/ghi/x5';
    
    var viewXml = {"ViewXml":
                        `<View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <Or>
                                        <Or>
                                            <BeginsWith>
                                                <FieldRef Name='FileDirRef' />
                                                <Value Type='Text'>${webUrl}/${documentLibrary}/${folder}</Value>
                                            </BeginsWith>
                                            <BeginsWith>
                                                <FieldRef Name='FileDirRef' />
                                                <Value Type='Text'>${webUrl}/${documentLibrary}/${folderB}</Value>
                                            </BeginsWith>
                                        </Or>
                                        <BeginsWith>
                                            <FieldRef Name='FileDirRef' />
                                            <Value Type='Text'>${webUrl}/${documentLibrary}/${folderC}</Value>
                                        </BeginsWith>
                                    </Or>
                                </Where>
                            </Query>
                        </View>`};
                        
    var data = await getDocumentLibrary_RecursiveAsync_inBody(webUrl, documentLibrary, folder, viewXml, fields);
    
    // inside FieldValuesAsText
    //    - FileLeafRef
    //    - FileRef
    //    - FileDirRef
    
    console.log(data);
})();

*/

// https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/upload-a-file-by-using-the-rest-api-and-jquery
// https://www.c-sharpcorner.com/blogs/upload-documents-show-documents-in-a-page-using-sharepoint-rest-api

// Get the local file as an array buffer.
function getFileBuffer(file) {
    let deferred = $.Deferred();
    
    var reader = new FileReader();
    
    reader.onloadend = function (e) {
        deferred.resolve(e.target.result);
    }
    reader.onerror = function (e) {
        deferred.reject(e.target.error);
    }
    
    reader.readAsArrayBuffer(file);
    
    return deferred.promise();
}

async function uploadFileToFolderAsync(webUrl, file, documentLibrary, folder){
    let retFunction = '';
    
    let deferred = $.Deferred();
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let arrayBuffer = getFileBuffer(file);
    
    let v_folder = '';
    if (folder){
        v_folder = `/${folder}`;
    }
    
    let targetUrl = `${webUrl}/${documentLibrary}${v_folder}`;
    let fileName = file.name;
    
    let url = webUrl + `/_api/Web/GetFolderByServerRelativeUrl(@target)/Files/add(overwrite=true, url='${fileName}')?@target='${targetUrl}'&$expand=ListItemAllFields`;
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: getFileBuffer,
        processData: false
    })
    .done(function(result) {
        try{
            retFunction = result.d.ServerRelativeUrl;
        }catch(e){}
        
        deferred.resolve(retFunction);
    })
    .fail(function(result, status) {
        //deferred.reject(status);
        
        // avoid: "Uncaught (in promise)"
        // resolve with return function: false
        deferred.resolve(retFunction);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
                        
    var data = await uploadFileToFolderAsync(webUrl, $("#txtFile")[0].files[0], documentLibrary);
    
    if (data){
        data = _spPageContextInfo.portalUrl + data;
        data = data.replace('.com//', '.com/');
    }
    
    console.log(data);
})();

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var documentLibrary = 'Documents';
    var folder = 'abc/def/ghi/x1';
                        
    var data = await uploadFileToFolderAsync(webUrl, $("#txtFile")[0].files[0], documentLibrary, folder);
    
    if (data){
        data = _spPageContextInfo.portalUrl + data;
        data = data.replace('.com//', '.com/');
    }
    
    console.log(data);
})();

// simple Function, calling await to Function_Async()
//
function uploadFile(){
    
    console.log('-- before async() statement');
    
    (async () => {
        console.log('Trying to upload file...');
        
        if (confirm("Read to upload file ?\nEither OK or Cancel.")){
            
            var webUrl = _spPageContextInfo.webServerRelativeUrl;
            
            var documentLibrary = 'Documents';
            var folder = 'abc/def/ghi/x1';
                                
            var data = await uploadFileToFolderAsync(webUrl, $("#txtFile")[0].files[0], documentLibrary, folder);
            
            if (data){
                data = _spPageContextInfo.portalUrl + data;
                data = data.replace('.com//', '.com/');
            }
            
            console.log(data);
        }
        
        console.log('Completed.');
    })();
    
    console.log('-- end');
}

*/

async function uploadFileToListItemAsync(webUrl, file, listName, id){
    let retFunction = '';
    
    let deferred = $.Deferred();
    
    let formDigestValue = await formDigestValueAsync(webUrl);
    
    let arrayBuffer = getFileBuffer(file);
    
    let fileName = file.name;
    
    let url = webUrl + `/_api/lists/GetByTitle('${listName}')/items(${id})/AttachmentFiles/add(FileName='${fileName}')`;
    
    $.ajax({
        url: url,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigestValue
        },
        data: getFileBuffer,
        processData: false
    })
    .done(function(result) {
        try{
            retFunction = result.d.ServerRelativeUrl;
        }catch(e){}
        
        deferred.resolve(retFunction);
    })
    .fail(function(result, status) {
        //deferred.reject(status);
        
        // avoid: "Uncaught (in promise)"
        // resolve with return function: false
        deferred.resolve(retFunction);
    });
    
    return deferred.promise();
}

/*
// Test:

// Immediately-invoked Function Expression (IIFE)

(async () => {
    var webUrl = _spPageContextInfo.webServerRelativeUrl;
    
    var listName = 'Person';
    var id = 97;
                        
    var data = await uploadFileToListItemAsync(webUrl, $("#txtFile")[0].files[0], listName, id);
    
    if (data){
        data = _spPageContextInfo.portalUrl + data;
        data = data.replace('.com//', '.com/');
    }
    
    console.log(data);
})();

*/
