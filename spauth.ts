import * as vscode from 'vscode';
import Window = vscode.window;
import * as https from 'https';
import * as constants from 'constants';
var cookie = require('cookie');
import * as fs from 'fs';
import spgit = require('./spgit');
import * as cp from 'child_process';

var Urls = {
    login: 'login.microsoftonline.com',
    signin: "/_forms/default.aspx?wa=wsignin1.0",
    sts: "/extSTS.srf"
};
var Error = function(e){
    console.log('Error', e);
};
var tokens = {
    security: '',
    access: ''
};

var credentials:Array<sp.Credential> = [];

var auth : sp.Auth;
var extensionPath:string;
var config:any;

var mkdir = (path:string, root?:string) => {
    var dirs = path.split('/'),
        dir = dirs.shift();
    root = (root || '') + dir + '/';
    try { fs.mkdirSync(root); }
    catch (e) {
        if(!fs.statSync(root).isDirectory()) throw new Error(e);
    }
    return !dirs.length || mkdir(dirs.join('/'), root);
};

module sp {
    export var getConfig = (path:string) => {
        extensionPath = path;
        config = JSON.parse(fs.readFileSync(path + '/spconfig.json', 'utf-8'));
        config.path += config.path.substring(config.path.length - 1, config.path.length) === '\\' ? '' : '\\';
    };
    // SharePoint authentication
    var authenticate = (error, data) => {
        var compiled:string = data.split('[username]').join(auth.project.user);
        compiled = compiled.split('[password]').join(auth.project.pwd);
        compiled = compiled.split('[endpoint]').join(auth.project.url);
        // 1. Send: XML with credentials, Get: Security token
        var getSecurityToken = new sp.Request();
        getSecurityToken.params.hostname = Urls.login;
        getSecurityToken.params.path = Urls.sts;
        getSecurityToken.params.method = 'POST';
        getSecurityToken.params.keepAlive = true;
        getSecurityToken.params.headers = {
            'Accept': 'application/json; odata=verbose',
            'Content-Type': 'application/xml',
            'Content-Length': Buffer.byteLength(compiled)
        };
        delete getSecurityToken.params.secureOptions;
        getSecurityToken.data = compiled;
        getSecurityToken.rawResult = true;
        getSecurityToken.send().then((data:string) => {
            var bits = data.split('<wsse:BinarySecurityToken Id="Compact0">');
            if(bits.length < 2) {
                Window.showErrorMessage('Authentication failed.');
                return false;
            }
            tokens.security = bits[1].split('</wsse:BinarySecurityToken>')[0];
            // 2. Send: Security token, Get: Access token (cookies)
            var getAccessToken = new sp.Request();
            getAccessToken.params.path = Urls.signin;
            getAccessToken.params.method = 'POST';
            getAccessToken.params.headers = {
                'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0)',
                'Content-Type': 'application/x-www-form-urlencoded',                
                'Content-Length': Buffer.byteLength(tokens.security)
            };
            getAccessToken.rawResult = true;
            getAccessToken.data = tokens.security;
            getAccessToken.onResponse = (res) => {
                var cookies = cookie.parse(res.headers["set-cookie"].join(";"));
                tokens.access =  'rtFa=' + cookies['rtFa'] + '; FedAuth=' + cookies['FedAuth'] + ';';
                auth.token = tokens.access;
            };
            getAccessToken.send().then(() => {
                mkdir(config.path + auth.project.title);
                spgit.init(config.path + auth.project.title, () => {
                    Window.showInformationMessage('GIT initialized');
                    var request = new sp.Request();
                    sp.get(config.folders, auth.project, tokens);
                });
            });
        });
    };
    // Init authentication
    export var open = (options:sp.Project) => {
		if( !options.title || !options.pwd || !options.url || !options.user) {
            Window.showWarningMessage('Please fill all the inputs');
            return false;
        }
        auth = new sp.Auth();
        storeCredentials(options);
        auth.project = options;
        fs.readFile(extensionPath + '/credentials.xml', 'utf-8', authenticate);
    };
    // Search for credentials used on a specific site
    export var suggestCredentials = (site:string) => {
        var suggestions:Array<sp.Credential> = credentials.filter((cred) => {
            return cred.site === site;
        });
        return suggestions;
    }
    export var checkFile = (file:string) => {
        var uptodate:boolean;
        var request = new sp.Request();
        request.params.path = '/_api/web/getfilebyserverrelativeurl(\'' + encodeURI(file) + '\')';
        return uptodate;
    }
    // Store credentials if don't exist yet
    var storeCredentials = (options:sp.Project) => {
        var exists:boolean = credentials.filter((cred) => {
            return cred.user === options.user;
        }).length > 0;
        if(exists) return false;
        credentials.push({'user':options.user,'password':options.pwd, 'site': options.url});
    };
    export interface Params {
        hostname: string;
        path?: string;
        method: string;
        secureOptions: number;
        headers: any;
        keepAlive?: boolean;
    }
    export interface Project {
        path?: string;
        title: string;
        url: string;
        user: string;
        pwd: string;
    }
    export interface Credential {
        user: string;
        password: string;
        site?: string;
    }
    export class Auth {
        token: string;
        digest: string;
        project: Project;
        constructor(){
        }
    }
    export class Request {
        digest: string;
        params: sp.Params;
        data: any;
        rawResult: boolean;
        onResponse: (any) => void;
        constructor () {
            this.params = {
                method: 'GET',
                hostname: auth.project.url.split('/')[2],
                secureOptions: constants.SSL_OP_NO_TLSv1_2,
                headers: {
                    'Accept': 'application/json; odata=nometadata'
                }
            };
            if (auth.token) this.params.headers.Cookie = auth.token;
            this.rawResult = false; 
        }
        send = () => {
            var self = this;
            var promise = new Promise((resolve,reject) => {
                if( !self.params.path.length ) {
                    console.warn('No request path specified.');
                    reject(null);
                    return false;
                }
                var request = https.request(self.params, (res) => {
                    if(self.onResponse) self.onResponse(res);
                    var data:string = '';
                    res.on('data', (chunk) => {
                        data += chunk;
                    });
                    res.on('error', (err) => {
                        console.warn('Request error:' + err);
                        reject(err);
                    });
                    res.on('end', () => {
                        var result = self.rawResult ? data : JSON.parse(data); 
                        resolve(result);
                    });
                });
                
                request.end(self.data || null);
            });
            return promise;
        }
    }
    export var get = (folders, project, tokens) => {
		if(!auth) {
            Window.showWarningMessage('You are not authenticated.');
            return false;
        }
		// 1. Get request digest
		var digest = new sp.Request();
        digest.params.path = '/_api/contextinfo';
        digest.params.method = 'POST';
        digest.send().then((data:any) => {
            auth.digest = data.FormDigestValue;
            var workfolder = config.path.split('\\').join('/') + auth.project.title;
            fs.writeFileSync(workfolder + '/spconfig.json', '{}');
            var promise = new Promise((resolve,reject) => {
                folders.forEach((folder, folderIndex) => {
                    // 2. Get list ID
                    var listId = new sp.Request();
                    listId.params.path = '/_api/web/GetFolderByServerRelativeUrl(\'' + encodeURI(folder) + '\')/properties?$select=vti_listname';
                    listId.send().then((data:any) => {
                        var id = data.vti_x005f_listname.split('{')[1].split('}')[0];
                        // 3. Get folder items
                        var listItems = new sp.Request();
                        listItems.params.path = '/_api/lists(\'' + id + '\')/getItems?$select=FileLeafRef,FileRef,FSObjType';
                        listItems.params.method = 'POST';
                        listItems.params.headers['X-RequestDigest'] = auth.digest;
                        listItems.params.headers['Content-Type'] = 'application/json; odata=verbose';
                        listItems.data = '{ "query" : {"__metadata": { "type": "SP.CamlQuery" }, "ViewXml": "<View Scope=\'RecursiveAll\'>';
                        listItems.data +=   '<Query><Where><And>';
                        listItems.data +=       '<Eq><FieldRef Name=\'FSObjType\' /><Value Type=\'Integer\'>0</Value></Eq>';
                        listItems.data +=       '<BeginsWith><FieldRef Name=\'FileRef\'/><Value Type=\'Text\'>' + folder + '</Value></BeginsWith>';
                        listItems.data +=   '</And></Where></Query>';  
                        listItems.data += '</View>"} }';
                        listItems.send().then((data:any) => {
                            var items = data.value;
                            Window.showInformationMessage(folder + ': downloading ' + items.length + ' items');
                            // 4. Download items, create folder structure if doesn't exist
                            items.forEach((item, itemIndex) => {
                                // TODO: Continue if should be ignored
                                mkdir(item.FileRef.split(item.FileLeafRef)[0], workfolder);
                                var download = new sp.Request();
                                download.rawResult = true;
                                download.params.path = '/_api/web/getfilebyserverrelativeurl(\'' + encodeURI(item.FileRef) + '\')/$value';
                                download.send().then((data:any) => {
                                    fs.writeFile(workfolder + item.FileRef, data, 'utf8', (err) => {
                                        if(itemIndex === items.length - 1 && folderIndex === folders.length - 1)
                                            resolve();
                                        if (err) throw err;
                                    });
                                });
                            });
                        });
                    });
                });
            });
            promise.then(() => {
                // Open code using the work folder
                
                // cp.exec('code ' + workfolder);
            });
        });
	}
};
export = sp;