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
				sp.open(project);
			});
		});
	});
	// Compare local and remote file dates and check check-out status
	var date = Commands.registerCommand('sp.date', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sbModified.show();
		sbModified.text = '$(sync) Checking file date';
		sbModified.tooltip = 'Comparing file dates between local and SharePoint';
		sbStatus.hide();
		sbStatus.text = '$(alert) Checked out';
		sbStatus.tooltip = 'File is presently checked out';
		sp.checkFileState(file).then((data:any) => {
			var modified:Date = new Date(data.TimeLastModified);
			var status:number = data.CheckOutType;
			if (!data.CheckOutType) sbStatus.show();
			sbModified.text = modified <= data.LocalModified ? '$(check) Fresh' : '$(alert) Old';
			sbModified.tooltip = modified <= data.LocalModified ? 'File is up to date or more recent' : 'File is not up to date';
		});
	});
	// Sync file
	var sync = Commands.registerCommand('sp.sync', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sp.download(file, vscode.workspace.rootPath).then((err:any) => {
			if (err) Window.showErrorMessage(err.message);
			else Window.showInformationMessage('File: ' + file.split('/').pop() + ' synced from SharePoint.');
		});
	});
	// Upload file
	var upload = Commands.registerCommand('sp.upload', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sp.upload(file).then((err:any) => {
			if (err) Window.showErrorMessage(err.message);
			else Window.showInformationMessage('File: ' + file.split('/').pop() + ' uploaded to SharePoint.');
		});
	});
	// Check in file
	var checkIn = Commands.registerCommand('sp.checkin', () => {
		
	});
	// Check out file
	var checkOut = Commands.registerCommand('sp.checkout', () => {
		
	});
	// Discard file checkout
	var discard = Commands.registerCommand('sp.discard', () => {
		
	});
	// Refresh workspace
	var refresh = Commands.registerCommand('sp.refresh', () => {
		
	});
	// Refresh workspace
	var resetCredentials = Commands.registerCommand('sp.credentials.reset', () => {
		Window.showWarningMessage('You are about to remove all your saved credentials.', 'Continue').then((selection) => {
			if (selection === 'Continue')
				context.globalState.update('sp.credentials', []).then(() => {
					Window.showInformationMessage('Credentials cache has been reset');
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
	
	context.subscriptions.push(connect, date, sync, upload, checkIn, checkOut, discard, refresh, resetCredentials);
}