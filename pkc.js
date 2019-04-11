var express = require('express');
var moment = require("moment");
var async = require('async');
var cron = require('node-cron');


var plm_request = require("request-promise-native").defaults({
    method: 'GET',
    json: true,
    // proxy:'http://web-proxy.corp.hpecorp.net:8088',
    headers: {
        "Access-Control-Allow-Origin": "*",
        "Accept": "application/json"
    }
});

const tenant_id = process.env.TENANT_ID ? process.env.TENANT_ID : "93f33571-550f-43cf-b09f-cd331338d086";
const client_id = process.env.CLIENT_ID ? process.env.CLIENT_ID : "efe8c10e-26c7-41c8-a3bd-cdc20be43acb";//45152e44-eb9d-400f-8d04-12a55744ed44";
const client_secret = process.env.CLIENT_SECRET ? process.env.CLIENT_SECRET : "V4WHshZJ5D6mTTJZuIAf526nlB9/OKPT26ZZBB6PrJ8=";//Gj9svXyShdU29FWvl0QlGN7tPymF1IJ9QuRKfT0ME3w=";
const resource = process.env.RESOURCE ? process.env.RESOURCE : "https://dxcportal.sharepoint.com";

refresh_token = process.env.FIRST_REFRESH_TOKEN ? process.env.FIRST_REFRESH_TOKEN : "AQABAAAAAADCoMpjJXrxTq9VG9te-7FX3GWItHtPqEpVcQ1G6-N5V24rvZ8QeY4532b7qTwHa73Jc_I4eP2coRJ1RX5__jF7cK3nsoVgo35FyT3K8OAl2Qjt3qJkd5VY3jx-1-yrPPG5Xab_BGP23RetKZkR6Sww423h7LhDL-5n9SMiP0GXeB9v9_gOtM258tbi2RsvIQo7xcazpbmPfCNiztTnoZelYAiHLopH24Fzw0oLcegjx2DgNQYFD2AZPfYZjv4j3K_L67BxuQkJWQaJLfmlvoU0papa8EQNVdLr2wxW-J4QwUib-Tm4_-l2lwSJj7_9dQvZWLVD3m8WQ10IDEQIRxBztQqCBzsrRFRUsclMyYeFZ-JJX0sNOClqufDs3MDCwB6gQq1WEqgQQ4IJ5CO6oTcL0STKmc8wlpISZjyk_0mN6WOGQUVPemYp7Jr5MWeSqmrRsTqa1IhdP-FroERmesZMAJ6WbJvl93Zdxehz_hZ9tOtR05jjpouAODTC5VztjHOCoyZlKINlBLBJqnmctb8Mzul531XFH-7sXjdRbNxH7llenAp6xo-yXeOCNiIIVxnqvlVz64wFGfZ32Qpgces3ynFbbta8Y-SKZPQGF6Xg3hnDqSltVoTireQN3tv7z8FOJDHLnseTkh_e7yJMPuDX4WByKlfxT2gBQMj4RR0EbtmtXgcXtH43JmycLpvfQ2h9HXQdwqQG4gD1Nt2WbNWWZ5JwxCYkuNnOnlh2ClvkoucF6XrDm8AoffAb0-NM_gQ8iFAEPFMALW2Y4dFuCDBBm-3FeK_jTCEP2BJawDKZyULEjI_OJsVOiXw8ixMpv0MBu1wb8ep_XgO145M3wli6BzFW-kLeYjO1g_roaKeBaSAA";
access_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyIsImtpZCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyJ9.eyJhdWQiOiJodHRwczovL2R4Y3BvcnRhbC5zaGFyZXBvaW50LmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzkzZjMzNTcxLTU1MGYtNDNjZi1iMDlmLWNkMzMxMzM4ZDA4Ni8iLCJpYXQiOjE1NTQ0NTE3MTMsIm5iZiI6MTU1NDQ1MTcxMywiZXhwIjoxNTU0NDU1NjEzLCJhY3IiOiIxIiwiYWlvIjoiNDJaZ1lKaTRicDdQWjQ0RmFrZWU3RE41VjkzRmxHY1dLT2oyZEVXbjg4RTB4eVFQdlZ3QSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiZ2V0QWNjZXNzVG9rZW4iLCJhcHBpZCI6ImVmZThjMTBlLTI2YzctNDFjOC1hM2JkLWNkYzIwYmU0M2FjYiIsImFwcGlkYWNyIjoiMSIsImZhbWlseV9uYW1lIjoiTE9QRVoiLCJnaXZlbl9uYW1lIjoiTUVSWU5FTExFIiwiaXBhZGRyIjoiMTkyLjQ2LjgyLjEwIiwibmFtZSI6IkxvcGV6LCBNZXJ5bmVsbGUiLCJvaWQiOiJjMmVlYzhiYS0yMzcwLTQzNWEtOGNhZi1lZTBjYjUwMDlkOTgiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjcxODcxMjg5My00MjU3ODkzMTAwLTM3ODY4MzU3Ni00MTgyMjMiLCJwdWlkIjoiMTAwMzNGRkZBRDBCQzI4MCIsInNjcCI6IkFsbFNpdGVzLk1hbmFnZSBBbGxTaXRlcy5SZWFkIEFsbFNpdGVzLldyaXRlIE15RmlsZXMuUmVhZCBNeUZpbGVzLldyaXRlIiwic2lkIjoiYTEwMWU3MzMtNzkxNC00MTEzLWFhZDMtMzk0NTE0ZTliMDI0Iiwic3ViIjoiYzRFOUQ1WlFHZURVVWhKX0VFUWRrc3NuM0lHNUhVM0oySzZQYlBzX29YUSIsInRpZCI6IjkzZjMzNTcxLTU1MGYtNDNjZi1iMDlmLWNkMzMxMzM4ZDA4NiIsInVuaXF1ZV9uYW1lIjoibWVyeW5lbGxlLmxvcGV6QGR4Yy5jb20iLCJ1cG4iOiJtZXJ5bmVsbGUubG9wZXpAZHhjLmNvbSIsInV0aSI6Im5mSGhZbTlWUVVPdHgySnctMndOQUEiLCJ2ZXIiOiIxLjAifQ.BZ7q2ylYYbEN4Rl5WPq-q8DXcIWdNv_jRKq825ozJzC_tDxLinRVROYZDsNdz9NgkAP6C_mB-OEkRpUPGDJpHxCNfPv00pbb2UT0a7jFMQaDK0EV90wScTzdT917-LUDfQ4MUf7VLV9w2pLIyk-1AZyWzfHSxmO1uyeyO3EWtP0AcBrIiSU0NM1x8dL0Q-Q7PIujaiK85i1nbgqEcx3OR0YRXcxLNxYG6jzR3eF6vYhnaOO5Eml8cXtlsFGe2v-utLMqpTw16B0H_LcjXSMLY5cS-H2Lk57-jkMYHApcQvqjwyiOLVoAS28Ik_fNy2ZSBKcEjW6V14tojUsoIzqt2w";

var refreshToken = function () {
    console.log('refreshing token');
    return plm_request({
        method: "POST",
        url: 'https://login.microsoftonline.com/' + tenant_id + '/oauth2/token',
        formData: {
            client_id: client_id,
            client_secret: client_secret,
            grant_type: 'refresh_token',
            refresh_token: refresh_token,
            resource: resource
        }
    }).then(body => {
        refresh_token = body.refresh_token;
        access_token = body.access_token

        console.log('successfully updated tokens');
    });
}

cron.schedule('*/30 * * * *', async function () {
    await refreshToken();
}).start();

var app = express();

app.use(function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
});

app.get('/token', function (req, res) {
    res.send({
        tenant_id: tenant_id,
        client_id: client_id,
        client_secret: client_secret,
        refresh_token: refresh_token,
        access_token: access_token
    });
});

app.get('/folders', async function (req, res) {
    var result = await getFoldersList();
    res.send(result);
});

app.get('/documents/:folder_id', async function (req, res) {
    var result = await getDocumentList(req.params.folder_id);
    res.send(result);
});

app.get('/documents/:folder_id/full', async function (req, res) {
    var result = await list(req.params.folder_id);
    console.log(result);
    res.send(result);
});

app.get('/destinationOffering', async function (req, res) {
    var result = await getDestinationOffering()
    res.send(result);
});

app.get('/process', async function (req, res) {
    var result = await processOfferingList()
    res.send(result);
});

app.get('/sync', async function (req, res) {
    var result = await sync()
    res.send(result);
});

function getFoldersList() {
    var folders = [];

    return plm_request({
        method: 'GET',
        url: resource + "/sites/omrep/_api/Web/GetFolderByServerRelativeUrl('Offerings%20Published')/Folders?$expand",
        headers: {
            Authorization: 'Bearer ' + access_token,
            Accept: 'application/json; odata=nometadata'
        }
    }).then(body => {
        folders = body.value.map(folder => {
            return {
                "folder_id": folder.Name,
                //"author_id": folder.AuthorId,
                "modified_date": folder.TimeLastModified,
                "created_date": folder.TimeCreated
            }
        });
        return folders;
    });
}

function getDocumentList(folder_id) {
    var documents = [];
    folder_id = "1543"
    return plm_request({
        method: 'GET',
        url: resource + "/sites/omrep/_api/Web/GetFolderByServerRelativeUrl('Offerings%20Published/" + folder_id + "')/Files",
        headers: {
            Authorization: 'Bearer ' + access_token,
            Accept: 'application/json; odata=nometadata'
        }
    }).then(body => {
        documents = body.value.map(document => {
            return {
                "Title": document.Name,
                "rel_url": encodeURI(document.ServerRelativeUrl),
                "FilePath": encodeURI(resource + "" + document.ServerRelativeUrl),
            }
        });
        return documents;

    });
}

async function getDocumentProperties(doc) {
    var document = doc;
    return plm_request({
        method: 'GET',
        url: resource + "/sites/omrep/_api/Web/GetFileByServerRelativePath(decodedurl='" + document.rel_url + "')/ListItemAllFields/FieldValuesAsText",
        headers: {
            Authorization: 'Bearer ' + access_token,
            Accept: 'application/json; odata=nometadata'
        }
    }).then(body => {
        document.VersionText = body.OData__x005f_UIVersionString;
        document.IdentifierText = body.UniqueId;
        document.Phase = body.Delivery_x005f_x0020_x005f_Phase;
        document.ImplementaionPhase = body.Implementation_x005f_x0020_x005f_Stream;

        //console.log(document);
        return document;
    });
}

async function list(folder_id) {
    var documents = await getDocumentList(folder_id);
    var list = Promise.all(documents.map(document => {
        return getDocumentProperties(document);
    })).then(results => {
        return results;
    })
    return list;
    
}

function getDestinationOffering() {
    subofferings = [];

    return plm_request({
        method: 'GET',
        url: 'https://dxcportal.sharepoint.com/sites/WM-ODT/Comms/VFTB/_api/web/lists/getByTitle%28%27ListSummary%27%29/items?$select=ID,VersionText,IdentifierText,Title,FilePath,Phase,ImplementaionPhase&$top=5000',//?$select=Title,TestColumn',
        headers: {
            Authorization: 'Bearer ' + access_token,
            Accept: 'application/json; odata=nometadata'
        }
    }).then(body => {
        subofferings = body.value.map(suboffering => {

            d_Phase = suboffering.Phase;
            d_ImplementaionPhase = suboffering.ImplementaionPhase

            if (d_Phase == null || d_ImplementaionPhase == null) {
                d_Phase = "n/a";
                d_ImplementaionPhase = "n/a";
                //console.log("Destination: NullError")
                return {
                    "ID": suboffering.ID,
                    "Title": suboffering.Title,
                    "IdentifierText": suboffering.IdentifierText,
                    "VersionText": suboffering.VersionText,
                    "FilePath": suboffering.FilePath,
                    "Phase": d_Phase,
                    //"ImplementaionPhase": suboffering.Implementation_x0020_Stream.trim()
                };
            }
            else {
                return {
                    "ID": suboffering.ID,
                    "Title": suboffering.Title,
                    "IdentifierText": suboffering.IdentifierText,
                    "VersionText": suboffering.VersionText,
                    "FilePath": suboffering.FilePath,
                    "Phase": suboffering.Phase,
                    "ImplementaionPhase": suboffering.ImplementaionPhase
                };
            }
        });
        //console.log('DESTINATION', subofferings);
        return subofferings;
    });
}

async function processOfferingList() {
    sourceOffering = await list();
    destinationOffering = await getDestinationOffering();

    to_add = [];
    to_update = [];
    to_delete = [];
    plm_ids = [];

    function checkOfferingList(subofferingsList) {

        var found = destinationOffering.some(function (destination) {
            return destination.IdentifierText == subofferingsList.IdentifierText;
        });

        //console.log("Source Identifier", subofferingsList.IdentifierText);

        if (!found) {
            //console.log("New item inserted", subofferingsList.Title);
            to_add.push({
                "Title": subofferingsList.Title,
                "IdentifierText": subofferingsList.IdentifierText,
                "VersionText": subofferingsList.VersionText,
                "FilePath": subofferingsList.FilePath,
                "Phase": subofferingsList.Phase,
                "ImplementaionPhase": subofferingsList.ImplementaionPhase,
            });

        }
        else {
            var indexOfferingList = destinationOffering.findIndex(des => des.IdentifierText == subofferingsList.IdentifierText);
            var toUpdateOfferingList = destinationOffering[indexOfferingList];

            toUpdateOfferingList.Title = subofferingsList.Title;
            toUpdateOfferingList.IdentifierText = subofferingsList.IdentifierText;
            toUpdateOfferingList.VersionText = subofferingsList.VersionText;
            toUpdateOfferingList.FilePath = subofferingsList.FilePath;
            toUpdateOfferingList.Phase = subofferingsList.Phase;
            toUpdateOfferingList.ImplementaionPhase = subofferingsList.ImplementaionPhase;

            to_update.push(toUpdateOfferingList);
            //console.log("Updated", toUpdateOfferingList);

        }

    }

    sourceOffering.map(function (arrayOfSource) {
        plm_ids.push(arrayOfSource.IdentifierText);
        checkOfferingList(arrayOfSource);
    });

    // var itemInSource = sourceOffering.filter(function (s) {
    //     return !destinationOffering.some(function (d) {
    //         return s.IdentifierText == d.IdentifierText;
    //     });
    // });

    // var itemDestination = destinationOffering.filter(function (d) {
    //     return !sourceOffering.some(function (s) {
    //         return d.IdentifierText == s.IdentifierText;
    //     });
    // });

    // to_delete = itemInSource.concat(itemDestination);
    // console.log("To delete", to_delete);

    return {
        "to_add": to_add,
        "to_update": to_update,
        "to_delete": to_delete,
    }
}


async function addSubOffering(suboffering) {
    var result;
    
    return plm_request({
        url: 'https://dxcportal.sharepoint.com/sites/WM-ODT/Comms/VFTB/_api/web/lists/getByTitle%28%27ListSummary%27%29/items?&%top=5000',
        method: 'POST',
        headers: {
            Authorization: 'Bearer ' + access_token,
            Accept: 'application/json; odata=nometadata',

        },
        body: {
            Title: suboffering.Title,
            IdentifierText: suboffering.IdentifierText,
            VersionText: suboffering.VersionText,
            FilePath: suboffering.FilePath,
            Phase: suboffering.Phase,
            ImplementaionPhase: suboffering.ImplementaionPhase
        }
    }).then(body => {
        result = {
            "Title" : body.result.Title,
            "IdentifierText": body.result.IdentifierText,
            "VersionText": body.result.VersionText,
            "FilePath": body.result.FilePath,
            "Phase": body.result.Phase,
            "ImplementaionPhase": body.result.ImplementaionPhase,
        };
        console.log('Add Offering', result);
        return result;
    });
}

async function updateSubOffering(suboffering) {
    console.log(suboffering.ID);

    var result;

    var updatedTitle = suboffering.Title;
    var updatedIdentifierText = suboffering.IdentifierText;
    var updatedVersionText = suboffering.VersionText;
    var updatedFilePath = suboffering.FilePath;
    var updatedPhase = suboffering.Phase;
    var updatedImplementaionPhase = suboffering.ImplementaionPhase;

    var result;

    return plm_request({
        url: 'https://dxcportal.sharepoint.com/sites/WM-ODT/Comms/VFTB/_api/web/lists/getByTitle%28%27ListSummary%27%29/GetItemById(' + suboffering.ID + ')?$select=ID,VersionText,IdentifierText,Title,FilePath,Phase,ImplementaionPhase',//?$filter=IdentifierText%20eq%20(%27' + suboffering.IdentifierText + '%27)',// + '%28' +suboffering.IdentifierText + '%29',
        method: 'PATCH',
        headers: {
            Authorization: 'Bearer ' + access_token,
            Accept: 'application/json; odata=nometadata',
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
            "content-type": "application/json;odata=verbose"
        },
        body: {
            __metadata: { type: 'SP.Data.List_x0020_SummaryListItem' },
            Title: updatedTitle,
            VersionText: updatedVersionText,
            FilePath: updatedFilePath,
            Phase: updatedPhase,
            ImplementaionPhase: updatedImplementaionPhase
        }
    }).then(body => {
        result = {
            "Title": updatedTitle,
            "VersionText": updatedVersionText,
            "FilePath": updatedFilePath,
            "Phase": updatedPhase,
            "ImplemtaionPhase": updatedImplementaionPhase
        }
        return result;
    });
}

// async function deleteSubOffering(suboffering) {
//     console.log(suboffering.ID);
//     var result;

//     return plm_request({
//         url: 'https://dxcportal.sharepoint.com/sites/WM-ODT/Comms/VFTB/_api/web/lists/getByTitle%28%27List%20Summary%27%29/items(' + suboffering.ID + ')?$select=ID,VersionText,IdentifierText,Title,FilePath,Phase,ImplementaionPhase&$top=5000',//?$filter=IdentifierText%20eq%20(%27' + suboffering.IdentifierText + '%27)',// + '%28' +suboffering.IdentifierText + '%29',
//         method: 'DELETE',
//         headers: {
//             Authorization: 'Bearer ' + access_token,
//             "IF-MATCH": "*",
//             "X-HTTP-Method-Override": "DELETE",    
//             "content-type": "application/json;odata=verbose"
//         },
//     }).then(body => {
//         result = "success" + body;
//         return result;
//     });
// }

async function sync() {
    result = await processOfferingList();

    async.map(result.to_add, suboffering => {
        addSubOffering(suboffering);
        console.log('check add suboffering', suboffering)
    }, (err, res) => {
        console.log(err ? "err:" + err : "res:" + res);
    });

    async.map(result.to_update, suboffering => {
        updateSubOffering(suboffering);
        //console.log('check suboffering', suboffering)
    }, (err, res) => {
        console.log(err ? "err:" + err : "res:" + res);
    });

    // async.map(result.to_delete, suboffering => {
    //     deleteSubOffering(suboffering);
    //     console.log('check deleted', suboffering)
    // }, (err, res) => {
    //     console.log(err ? "err:" + err : "res:" + res);
    // });
}

// cron.schedule('0 0 */14 * SAT', async function(){
//     await sync();
//   }).start();

app.listen(3000, function () {
    console.log('Example app listening on port 3000!');
    refreshToken();
}); 