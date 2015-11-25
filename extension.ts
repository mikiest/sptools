import * as vscode from 'vscode';
import Window = vscode.window;
import Commands = vscode.commands;
import sp = require('./spcore');
import helpers = require('./helpers');
import * as fs from 'fs';
import * as cp from 'child_process';

export function activate(context: vscode.ExtensionContext) {
	// Store extension context for SP module
	sp.getContext(context);
	// Status bar icons
	var sbModified:vscode.StatusBarItem = Window.createStatusBarItem(vscode.StatusBarAlignment.Left);
	var sbStatus:vscode.StatusBarItem = Window.createStatusBarItem(vscode.StatusBarAlignment.Left);
	sbModified.command = 'sp.date';
	sbStatus.command = 'sp.checkinout';
	// Update status bar indicators
	var updateStatus = (data:any) => {
		var status:number = data.CheckOutType;
		sbStatus.hide();
		sbStatus.text = '$(link-external) Checked out';
		sbStatus.tooltip = 'File is presently checked out';
		if (!status) {
			sbStatus.tooltip += ' to ' + (data.CheckedOutByMe ? 'you' : data.CheckedOutBy.Title);
			if (data.CheckedOutByMe) sbStatus.text = '$(link-external) Checked out to you';
		} else {
			sbStatus.text = '$(check) Checked in';
			sbStatus.tooltip = 'File is presently checked in';
		}
		sbStatus.show();
	};
	// Wether the opened file should be ignored
	var ignoreFile = (file:string) => {
		var length:number = file.length;
		var ignore:boolean = file === '/.vscodeignore'
			|| file === '/.gitignore'
			|| file === '/spconfig.json'
			|| file.substring(0, 8) === '/.vscode'
			|| file.substring(0, 5) === '/.git';
		try { fs.statSync(vscode.workspace.rootPath + '/spconfig.json'); }
		catch (e){ ignore = true;}
		return !vscode.workspace.rootPath || ignore;
	};
	
	// Init new SP workspace
	var connect = Commands.registerCommand('sp.connect', () => {
		sp.getConfig(context.extensionPath).then(() => {
			var project = <sp.Project>{};
			var options: vscode.InputBoxOptions = {
				prompt: 'Project name?',
				placeHolder: 'Project name'
			};
			// Ask for workspace params: SP URL, project title, credentials
			Window.showInputBox(options).then((selection) => {
				project.title = selection;
				options.prompt = 'Site URL?';
				options.placeHolder = 'http(s)://domain.com';
				Window.showInputBox(options).then((selection) => {
					project.url = selection;
					sp.open(project).then(() => {
						Window.showInformationMessage('Workspace created.', 'Open').then((selection) => {
							var workFolder:string = vscode.workspace.getConfiguration('sptools').get<string>('workFolder');
							if (workFolder === '$home') workFolder = (process.platform === 'win32' ? process.env.HOMEPATH : process.env.HOME) + '\\sptools';
							if (selection === 'Open')
								cp.exec('code "' + workFolder + '/' + project.title + '"', function (error, stdout, stderr) {
									console.log('stdout: ' + stdout);
									console.log('stderr: ' + stderr);
									if (error !== null) {
									console.log('exec error: ' + error);
									}
								});
						});
						// TODO: Suggest to open workspace
					});
				});
			});
		}, (message:any) => {
			Window.showWarningMessage(message);
		});
	});
	// Compare local and remote file dates and check check-out status
	var date = Commands.registerCommand('sp.date', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		if (ignoreFile(file)) return false;
		sbModified.show();
		sbModified.text = '$(sync) Checking file state';
		sbModified.tooltip = 'Comparing file dates between local and SharePoint';
		sp.checkFileState(file).then((data:any) => {
			updateStatus(data);
			if (data.UpToDate) {
				sbModified.text = '$(check) Up to date';
				sbModified.command = 'sp.date';
				sbModified.tooltip = 'File is up to date or more recent';
			} else {
				sbModified.text = '$(alert) Update required';
				sbModified.command = 'sp.sync';
				sbModified.tooltip = 'File is not up to date';
			}
		});
	});
	// Sync file
	var sync = Commands.registerCommand('sp.sync', () => {
		var file:string = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		var fileLeaf:string = file.split('/').pop();
		sp.checkFileState(file).then((props:any) => {
			// Check dates
			var promise = new Promise((resolve, reject) => {
				if (!props.UpToDate || props.TimeLastModified.getTime() === props.LocalModified.getTime()) {
					resolve();
					return false;
				}
				Window.showWarningMessage(fileLeaf + ' is older on server. Local changes might be lost if you continue.', 'Continue').then((selection) => {
					if (selection === 'Continue') {
						resolve();
						return false;
					}
				})
			});
			promise.then(() => {
				// Download
				sp.download(file, vscode.workspace.rootPath).then((err:any) => {
					if (err) Window.showErrorMessage(err.message);
					else {
						Window.showInformationMessage(fileLeaf + ' synced from SharePoint.');
						vscode.commands.executeCommand('sp.date');
					}
				});
			});
		});
	});
	// Upload file
	var upload = Commands.registerCommand('sp.upload', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		var fileLeaf:string = file.split('/').pop();
		sp.checkFileState(file).then((props:any) => {
			// Check dates
			var promise = new Promise((resolve, reject) => {
				if (props.UpToDate) {
					resolve();
					return false;
				}
				Window.showWarningMessage(fileLeaf + ' is newer on server. Changes on the site might be lost if you continue.', 'Continue').then((selection) => {
					if (selection === 'Continue')
						resolve();
				})
			});
			promise.then(() => {
				// Check check out status
				var promise = new Promise((resolve,reject) => {
					if (props.CheckedOutByMe)
						resolve();
					else
						// " is checked out" or " is not checked out"
						Window.showWarningMessage(fileLeaf + ' is ' 
							+ (!props.CheckOutType ? '' : 'not')
							+ ' checked out'
							+ (!props.CheckOutType ? ' to ' + props.CheckedOutBy.Title : '')
							+ '.',
						// "Discard, checkout and upload" or "Check out and upload"
							(!props.CheckOutType ? 'Discard, c' : 'C')
							+ 'heck out and upload'
						).then((selection) => {
							var action:number = !props.CheckOutType ? 2 : 1;
							if (selection && selection.length)
								sp.checkinout(file, action).then(() => {
									if (!props.CheckOutType)
										sp.checkinout(file, 1).then(() => {
											resolve();
										});
									else
										resolve();
								});
						});
				});
				promise.then(() => {
					// Upload
					sp.upload(file).then((err:any) => {
						if (err) Window.showErrorMessage(err.message);
						else {
							Window.showInformationMessage(fileLeaf + ' uploaded to SharePoint, checked out to you.', 'Check in').then((selection) => {
								vscode.commands.executeCommand('sp.date');
								var promise = new Promise((resolve,reject) => {
									if (selection === 'Check in')
										sp.checkinout(file, 0).then(() => {
											resolve();
										});
									else 
										resolve();
								});
								promise.then(() => {
									sp.checkFileState(file).then((newProps:any) => {
										var modified:number = new Date(newProps.TimeLastModified).getTime() / 1000 | 0;
										fs.utimes(vscode.workspace.rootPath + file, modified, modified, (err) => {
											vscode.commands.executeCommand('sp.date');
											if (selection === 'Check in')
												Window.showInformationMessage(fileLeaf + ' is now checked in.');
											if (err) throw err;
										});
									});
								});
							});
						}
					});
				});
			});
		});
	});
	// Check in, out or discard current file checkout
	var checkInOut = Commands.registerCommand('sp.checkinout', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sp.checkFileState(file).then((props:any) => {
			var status:number = props.CheckOutType;
			var checkinLabel:string = 'Check in';
			var checkoutLabel:string = 'Check out';
			var discardLabel:string = 'Discard checkout';
			var continueLabel:string = 'Continue';
			var fileLeaf:string = file.split('/').pop();
			updateStatus(props);
			var success = (data:any, message) => {
				updateStatus(data);
				Window.showInformationMessage(fileLeaf + message);
			};
			// File is checked out
			if (!status) {
				// File is checked out to current user
				if (props.CheckedOutByMe)
					Window.showInformationMessage(fileLeaf + ' is checked out to you.', checkinLabel, discardLabel).then((selection) => {
						if (selection === checkinLabel) {
							var promise = new Promise((resolve, reject) => {
								if (props.TimeLastModified.getTime() === props.LocalModified.getTime()) resolve();
								else
									Window.showWarningMessage(fileLeaf + ' is '+ (props.UpToDate ? 'older' : 'more recent') + ' on server.' , 'Keep server version', 'Upload local version').then((selection) => {
										var action:string = selection === 'Keep server version' ? 'download': 'upload';
										sp[action](file).then((data:any) => {
											resolve();
										});
									});
							});
							promise.then(() => {
								sp.checkinout(file, 0).then((data:any) => {
									sp.checkFileState(file).then((newProps:any) => {
										var modified:number = new Date(newProps.TimeLastModified).getTime() / 1000 | 0;
										fs.utimes(vscode.workspace.rootPath + file, modified, modified, (err) => {
											success(newProps, ' is now checked in.');
											if (err) throw err;
										});
									});
								});
							}, () => {}); 
						}
						else if (selection === discardLabel)
							Window.showWarningMessage('You are about to discard your changes on ' + fileLeaf, 'Continue').then((selection) => {
								sp.checkinout(file, 2).then(() => {
									sp.checkFileState(file).then((data:any)=> {
										success(data, ' checkout has been discarded.');
									});
								});
							});
					});
				// File is checked out to another user
				else
					Window.showInformationMessage(fileLeaf + ' is checked out to ' + props.CheckedOutBy.Title, discardLabel).then((selection) => {
						if (selection === discardLabel)
							Window.showWarningMessage('You are about to discard the changes by ' + props.CheckedOutBy.Title + ' on ' + fileLeaf, 'Continue').then((selection) => {
								sp.checkinout(file, 2).then(() => {
									sp.checkFileState(file).then((data:any)=> {
										success(data, ' checkout has been discarded.');
									});
								});
							});
					});
			}
			// File is not checked out
			else
				Window.showInformationMessage(fileLeaf + ' is not checked out', checkoutLabel).then((selection) => {
					if (selection === checkoutLabel)
						sp.checkinout(file, 1).then(() => {
							sp.checkFileState(file).then((data:any)=> {
								success(data, ' is now checked out to you.');
							});
						});
				});
		});
	});
	// Refresh workspace
	var refresh = Commands.registerCommand('sp.refresh', () => {
		Window.showInformationMessage('Command not implemented yet.');
	});
	// Refresh workspace
	var resetCredentials = Commands.registerCommand('sp.credentials.reset', () => {
		Window.showWarningMessage('You are about to remove all your saved credentials.', 'Continue').then((selection) => {
			if (selection === 'Continue')
				context.globalState.update('sp.credentials', []).then(() => {
					Window.showInformationMessage('Credentials cache has been reset. Existing contexts will persist until Code has been restarted.');
				});
		});
	});
	var openLimiter;
	// Check file sync
	vscode.workspace.onDidOpenTextDocument((e) => {
		clearTimeout(openLimiter);
		openLimiter = null;
		openLimiter = setTimeout(() => {
			vscode.commands.executeCommand('sp.date');
		}, 500);
	}, this, context.subscriptions);

	context.subscriptions.push(connect, date, sync, upload, checkInOut, refresh, resetCredentials);
}