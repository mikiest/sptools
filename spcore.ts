import * as vscode from 'vscode';
import Window = vscode.window;
import * as https from 'https';
import * as http from 'http';
import * as constants from 'constants';
var cookie = require('cookie');
import * as fs from 'fs';
import * as cp from 'child_process';
import helpers = require('./helpers');
var httpntlm = require('httpntlm');

var Urls = {
    login: 'login.microsoftonline.com',
    signin: "/_forms/default.aspx?wa=wsignin1.0",
    sts: "/extSTS.srf"
};
var tokens = {
    security: '',
    access: ''
};

var version:number;
var jsonLight:boolean = true;

var auth:sp.Auth;
var config;
var wkConfig:any;
var ctx:vscode.ExtensionContext;

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
    export interface Params {
        hostname: string;
        path?: string;
        method?: string;
        secureOptions?: number;
        headers: any;
        keepAlive?: boolean;
    }
    export interface Project {
        site?: string;
        title: string;
        url: string;
        user: string;
        pwd: string;
    }
    export class Auth {
        token: string;
        digest: string;
        project: Project;
        credentials: helpers.spCredentials;
        constructor(){
        }
    }
    // Request wrapper
    export class Request {
        digest: string;
        ssl: boolean;
        params: sp.Params;
        data: any;
        rawResult: boolean;
        ignoreAuth: boolean;
        onResponse: (any) => void;
        constructor () {
            this.ssl = auth.project.url.split(':')[0] === 'https';
            this.params = {
                method: 'GET',
                hostname: auth.project.url.split('/')[2],
                headers: {
                    'Accept': 'application/json; odata=' + (jsonLight ? 'nometadata' : 'verbose')
                }
            };
            if (this.ssl) this.params.headers.secureOptions = constants.SSL_OP_NO_TLSv1_2;
            if (auth.token && version === 16) this.params.headers.Cookie = auth.token;
            this.rawResult = false; 
        }
        private error = (error:string) => {
            Window.showWarningMessage(error || 'Sorry, something went wrong.');
        }
        // Send and authenticate if needed
        send = () => {
            var self = this;
            var authenticated = new Promise((resolve, reject) => {
                if (auth.token || self.ignoreAuth || auth.credentials) resolve();
                else
                    authenticate().then(() => {
                        resolve();
                    });
            });
            var promise = new Promise((resolve,reject) => {
                var prem = () => {
                    var options:any = {
                        url: auth.project.url + self.params.path,
                        username: auth.credentials.username,
                        password: auth.credentials.password,
                        workstation: '',
                        domain: '',
                        headers: self.params.headers
                    };
                    if (options.headers.Accept === "application/json; odata=nometadata" && !jsonLight) options.headers.Accept = "application/json; odata=verbose";
                    if (self.data) options.body = self.data;
                    if (self.params.headers['X-RequestDigest']) options.headers['X-RequestDigest'] = self.params.headers['X-RequestDigest'];
                    httpntlm[self.params.method.toLocaleLowerCase()](options, (err, res) => {
                        if (err) {
                            this.error(err.message);
                            return false;
                        }
                        if (res.statusCode === 401) {
                            self.error('Authentication failed.');
                            return false;
                        }
                        else if (res.statusCode === 500) {
                            self.error(JSON.parse(res.body).error.message);
                            return false;
                        }
                        var result = self.rawResult ? res.body : JSON.parse(res.body); 
                        resolve(result);
                    });
                };
                authenticated.then(() => {
                    if (!self.ignoreAuth) self.params.path = auth.project.site + self.params.path;
                    if (!self.params.headers.Cookie && auth.token && version === 16) self.params.headers.Cookie = auth.token;
                    if( !self.params.path.length ) {
                        console.warn('No request path specified.');
                        // reject(null);
                        return false;
                    }
                    var needsDigest = new Promise((yes,no) => {
                        if (!self.ignoreAuth && self.params.method === 'POST' && self.params.path.substring(self.params.path.length - 16, self.params.path.length) !== '_api/contextinfo') {
                            var digest = new sp.Request();
                            digest.params.path = '/_api/contextinfo';
                            digest.params.method = 'POST';
                            digest.send().then((data:any) => {
                                var digest = data.FormDigestValue || data.d.GetContextWebInformation.FormDigestValue;
                                auth.digest = digest;
                                self.params.headers['X-RequestDigest'] = digest;
                                yes();
                            });
                        }
                        else yes();
                    });
                    needsDigest.then(() => {
                        var req = self.ssl ? https.request : http.request;
                        if (version !== 16) {
                            prem();
                            return false;
                        }
                        var request = req(self.params, (res) => {
                            if (self.onResponse) self.onResponse(res);
                            if (version !== 16) return false;
                            var data:string = '';
                            res.on('data', (chunk) => {
                                data += chunk;
                            });
                            res.on('error', (err) => {
                                // reject();
                                self.error(err.message);
                            });
                            res.on('end', () => {
                                var result = self.rawResult ? data : (jsonLight ? JSON.parse(data) : JSON.parse(data).d); 
                                if (!self.rawResult && result['odata.error']) {
                                    self.error(result['odata.error'].message.value);
                                    // reject();
                                    return false;
                                }
                                resolve(result);
                            });
                        });
                        request.end(self.data || null);
                    });
                });
            });
            return promise;
        }
    }
    // Parse URL and get site collection URL
    var getSiteCollection = (url:string) => {
        var last:string = url[url.length - 1];
        if (last === '/') url = url.substring(0, url.length - 1);
        var split = url.split('/');
        var domain:string = split[2];
        return (split.length > 3) ? url.split(domain)[1] : '';
    };
    // Get nometadata compatibility
    var metadata = (url:string) => {
        var promise = new Promise((resolve,reject) => {
            var req = new sp.Request();
            req.params.path = '/_api/web';
            req.params.headers.Accept = 'application/json; odata=nometadata';
            req.rawResult = true;
            req.send().then((data:any) => {
                try {
                    JSON.parse(data);
                } catch (e) {
                    jsonLight = false;
                }
                resolve();
            });
        });
        return promise;
    };
    // SharePoint authentication
    var authenticate = () => {
        // DOING: Check online/prem/2010 and behave differently
        var get = auth.project.url.split(':')[0] === 'https' ? https.get : http.get;
        
        var credentials = new helpers.Credentials();
        var promise = new Promise((resolve, reject) => {
            get(auth.project.url, (res) => {
                if (res.statusCode === 400 || !res.headers.microsoftsharepointteamservices) {
                    Window.showErrorMessage('Please enter a valid URL');
                    return;
                }
                version = parseInt(res.headers.microsoftsharepointteamservices.split('.')[0]);
                credentials.get(auth.project.url.split('/')[2]).then((credentials:helpers.spCredentials) => {
                    if (version !== 16) {
                        auth.credentials = credentials;
                        metadata(auth.project.url).then(() => {
                            sp.getCurrentUserProperties().then(() => {
                                resolve();
                            })
                        });
                        return false;
                    }
                    var enveloppe:string = fs.readFileSync(ctx.extensionPath + '/credentials.xml', 'utf8');
                    var compiled:string = enveloppe.split('[username]').join(credentials.username);
                    compiled = compiled.split('[password]').join(credentials.password);
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
                    getSecurityToken.ignoreAuth = true;
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
                        getAccessToken.ignoreAuth = true;
                        getAccessToken.onResponse = (res) => {
                            var cookies = cookie.parse(res.headers["set-cookie"].join(";"));
                            tokens.access =  'rtFa=' + cookies['rtFa'] + '; FedAuth=' + cookies['FedAuth'] + ';';
                            auth.token = tokens.access;
                        };
                        getAccessToken.send().then(() => {
                            sp.getCurrentUserProperties().then(() => {
                                resolve();
                            })
                        });
                    });
                });
            }).on('error', (e) => {
                Window.showErrorMessage('Error contacting ' + auth.project.url.split('/')[2] + '. Please check the URL or your network.');
            });            
        });
        return promise;
    };
    // Get file properties
    var getProperties = (fileName:string) => {
        var promise = new Promise((resolve, reject) => {
            var properties = new sp.Request();
            properties.params.path = '/_api/web/getfilebyserverrelativeurl(\'' + encodeURI(fileName) + '\')/?$select=Name,ServerRelativeUrl,CheckOutType,TimeLastModified';
            properties.send().then((data:any) => {
                var result = data.d || data;
                if (result.CheckOutType !== 0) resolve(result);
                else {
                    var checkedOutBy = new sp.Request();
                    checkedOutBy.params.path = '/_api/web/getfilebyserverrelativeurl(\'' + encodeURI(fileName) + '\')/Checkedoutbyuser?$select=Title,Email';
                    checkedOutBy.send().then((user:any) => {
                        result.CheckedOutBy = user.d || user;
                        resolve(result)
                    });
                }
            });
        });
        return promise;
    }
    // Init workspace
    export var open = (options:sp.Project) => {
		if( !options.title || !options.url) {
            Window.showWarningMessage('Please fill all the inputs');
            return;
        }
        var workfolder = config.path + options.title;
        mkdir(workfolder);
        fs.writeFileSync(workfolder + '/spconfig.json', '{"site": "' + options.url + '"}');
        auth.project = options;
        auth.project.site = getSiteCollection(options.url);
        var request = new sp.Request();
        return sp.get(config.spFolders, tokens);
    };
    // Get and store Extension context
    export var getContext = (context:vscode.ExtensionContext) => {
        auth = new sp.Auth();
        auth.project = <sp.Project>{};
        fs.readFile(vscode.workspace.rootPath + '/spconfig.json', 'utf-8', (err, data:string) => {
            if (err) return;
            wkConfig = JSON.parse(data);
            auth.project.url = wkConfig.site;
            auth.project.site = getSiteCollection(wkConfig.site);
        });
        helpers.getContext(context);
        ctx = context;
    };
    // Get Extension settings
    export var getConfig = (path:string) => {
        var promise = new Promise((resolve,reject) => {
            config = vscode.workspace.getConfiguration('sptools');
            var wk:string = config.workFolder;
            var isWin:boolean = process.platform === 'win32';
            if (wk === '$home') wk = (isWin ? process.env.HOMEPATH : process.env.HOME) + '\\sptools';
            config.path = wk + (wk.substring(wk.length - 1, wk.length) === '\\' ? '' : '\\');
            if (!isWin) config.path = config.path.split('\\').join('/');
            try {
                fs.statSync(config.path);
                resolve();
            }
            catch (err) {
                Window.showWarningMessage(config.path + ' ("workFolder" setting) does not exist.', 'Create').then((selection) => {
                    if (selection === 'Create') {
                        mkdir(config.path);
                        resolve();
                    } else
                        reject('Please check your "workFolder" setting.');
                });
            }
        });
        return promise;
        
    };
    // Check file dates and status
    export var checkFileState = (file:string) => {
        var promise = new Promise((resolve, reject) => {
            getProperties(file).then((data:any) => {
                fs.stat(vscode.workspace.rootPath + file, (err, stats) => {
                    var local:Date = stats.mtime;
                    data.LocalModified = local;
                    resolve(data);
                });
            });
        });
        return promise;
    }
    // Resolve and download files
    export var get = (folders, tokens) => {
		var workfolder = config.path.split('\\').join('/') + auth.project.title;
        var promise = new Promise((resolve,reject) => {
            authenticate().then(() => {
                var count:number = 0;
                var notFound:number = 0;
                folders.forEach((folder, folderIndex) => {
                    // 1. Get list ID
                    var listId = new sp.Request();
                    listId.params.path = '/_api/web/GetFolderByServerRelativeUrl(\'' + encodeURI(auth.project.site + folder) + '\')/properties?$select=vti_listname';
                    listId.send().then((data:any) => {
                        var error = data.error;
                        if (error) {
                            if (error.message === 'File Not Found.' || error.message.value === 'File Not Found.') {
                                notFound++;
                                Window.showWarningMessage(folder + ' not found.');
                                if (notFound === folders.length)
                                    Window.showErrorMessage('Nothing to fetch. Please check your "spFolders" setting.')
                            }
                            else Window.showWarningMessage(error.message.value || error.message);
                            return false;
                        }
                        var id = (data.d || data).vti_x005f_listname.split('{')[1].split('}')[0];
                        // 2. Get folder items
                        var listItems = new sp.Request();
                        listItems.params.path = '/_api/lists(\'' + id + '\')/getItems?$select=FileLeafRef,FileRef,FSObjType,Modified';
                        listItems.params.method = 'POST';
                        listItems.params.headers['Content-Type'] = 'application/json; odata=verbose';
                        listItems.data = '{ "query" : {"__metadata": { "type": "SP.CamlQuery" }, "ViewXml": "<View Scope=\'RecursiveAll\'>';
                        listItems.data +=   '<Query><Where><And>';
                        listItems.data +=       '<Eq><FieldRef Name=\'FSObjType\' /><Value Type=\'Integer\'>0</Value></Eq>';
                        listItems.data +=       '<BeginsWith><FieldRef Name=\'FileRef\'/><Value Type=\'Text\'>' + auth.project.site + folder + '</Value></BeginsWith>';
                        listItems.data +=   '</And></Where></Query>';
                        listItems.data += '</View>"} }';
                        listItems.send().then((data:any) => {
                            var items = data.value || data.d.results;
                            count += items.length;
                            if (folderIndex === folders.length - 1)
                                Window.showInformationMessage('Fetching ' + count + ' items from ' + auth.project.url + '.');
                            // 3. Download items, create folder structure if doesn't exist
                            items.forEach((item, itemIndex) => {
                                // TODO: Continue if should be ignored
                                mkdir(item.FileRef.split(item.FileLeafRef)[0], workfolder);
                                sp.download(item.FileRef, workfolder).then(() => {
                                    if(itemIndex === items.length - 1 && folderIndex === folders.length - 1)
                                        resolve();
                                });
                            });
                        });
                    });
                });
            });
            
        });
        return promise;
	}
    // Download specific file
    export var download = (fileName:string, workfolder:string) => {
        var promise = new Promise((resolve,reject) => {
            getProperties(fileName).then((props:any) => {
                var download = new sp.Request();
                download.rawResult = true;
                download.params.path = '/_api/web/getfilebyserverrelativeurl(\'' + encodeURI(fileName) + '\')/$value';
                download.send().then((data:any) => {
                    fs.writeFile(workfolder + fileName, data, 'utf8', (err) => {
                        var modified:number = new Date(props.TimeLastModified).getTime() / 1000 | 0;
                        fs.utimes(workfolder + fileName, modified, modified, (err) => {
                            resolve(err);
                            if (err) throw err;
                        })
                        if (err) throw err;
                    });
                });
            });
        });
        return promise;
    }
    // Resolve and download files
    export var upload = (fileName:string) => {
        var fileLeaf:string = fileName.split('/').pop();
        var folder:string = fileName.split(fileLeaf)[0];
        var promise = new Promise((resolve, reject) => {
            fs.readFile(vscode.workspace.rootPath + fileName, (err, data) => {
                if (err) throw err;
                var upload = new sp.Request();
                upload.params.path = '/_api/web/getfolderbyserverrelativeurl(\'' + encodeURI(folder) + '\')/files/add(overwrite=true,url=\'' + encodeURI(fileLeaf) + '\')';
                upload.params.method = 'POST';
                upload.data = data;
                upload.send().then(() => {
                    resolve();
                });
            });
        });
        return promise;
    };
    // Check in, out or discard checkout
    export var checkinout = (file:string, action:number) => {
        var suffix:string;
        var success:string;
        var request = new sp.Request()
        if (!action) {
            suffix = 'CheckIn(' + encodeURI('comment=\'' + vscode.workspace.getConfiguration('sptools').get('checkInComment') + '\',checkintype=1') + ')';
            success = ' is now checked in.';
        }
        else if (action === 1) {
            suffix = 'checkout';
            success = ' is now checked out to you.';
        }
        else if (action === 2) {
            suffix = 'undocheckout';
            success = ' check out has been discarded.';
        }
        request.params.path = '/_api/web/GetFileByServerRelativeUrl(\'' + encodeURI(file) + '\')/' + suffix;
        request.params.method = 'POST';
        return request.send();
    }
    export var getCurrentUserProperties = () => {
        var user = new sp.Request();
        user.params.path = '/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=DisplayName,Email';
        return user.send().then((data:any) => {
            helpers.setCurrentUser(data.d || data);
        });
    };
};
export = sp;