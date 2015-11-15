import * as vscode from 'vscode';
import Window = vscode.window;
import Commands = vscode.commands;
import sp = require('./spcore');
import helpers = require('./helpers');
import * as fs from 'fs';

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
			sbStatus.tooltip += ' to ' + ((helpers.currentUser.email === data.CheckedOutBy.Email) ? 'you' : data.CheckedOutBy.Title);
			if (helpers.currentUser.email === data.CheckedOutBy.Email) sbStatus.text = '$(link-external) Checked out to you';
			sbStatus.show();
		}
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
		sp.getConfig(context.extensionPath);
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
					Window.showInformationMessage('Workspace created.');
					// TODO: Suggest to open workspace
				});
			});
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
			var modified:Date = new Date(data.TimeLastModified);
			updateStatus(data);
			sbModified.text = modified <= data.LocalModified ? '$(check) Up to date' : '$(alert) Update required';
			sbModified.tooltip = modified <= data.LocalModified ? 'File is up to date or more recent' : 'File is not up to date';
		});
	});
	// Sync file
	var sync = Commands.registerCommand('sp.sync', () => {
		// TODO check file status
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sp.download(file, vscode.workspace.rootPath).then((err:any) => {
			if (err) Window.showErrorMessage(err.message);
			else Window.showInformationMessage('File: ' + file.split('/').pop() + ' synced from SharePoint.');
		});
	});
	// Upload file
	var upload = Commands.registerCommand('sp.upload', () => {
		// TODO check file status
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sp.upload(file).then((err:any) => {
			if (err) Window.showErrorMessage(err.message);
			else Window.showInformationMessage('File: ' + file.split('/').pop() + ' uploaded to SharePoint.');
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
				if (helpers.currentUser.email === props.CheckedOutBy.Email)
					Window.showInformationMessage(fileLeaf + ' is checked out to you.', checkinLabel, discardLabel).then((selection) => {
						if (selection === checkinLabel) {
							var modified:Date = new Date(props.TimeLastModified);
							var uptodate:boolean = modified <= props.LocalModified;
							var promise = new Promise((resolve, reject) => {
								if (modified.getTime() === props.LocalModified.getTime()) resolve();
								else
									Window.showWarningMessage(fileLeaf + ' is '+ (uptodate ? 'older' : 'more recent') + ' on server.' , 'Keep server version', 'Upload local version').then((selection) => {
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
		
	});
	// Refresh workspace
	var resetCredentials = Commands.registerCommand('sp.credentials.reset', () => {
		Window.showWarningMessage('You are about to remove all your saved credentials.', 'Continue').then((selection) => {
			if (selection === 'Continue')
				context.globalState.update('sp.credentials', []).then(() => {
					Window.showInformationMessage('Credentials cache has been reset.');
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