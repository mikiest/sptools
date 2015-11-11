import * as vscode from 'vscode';
import Window = vscode.window;
import Commands = vscode.commands;
import sp = require('./spcore');
import helpers = require('./helpers');
import * as fs from 'fs';


export function activate(context: vscode.ExtensionContext) {
	
	sp.getContext(context);
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
	var date = Commands.registerCommand('sp.date', () => {
		var file = Window.activeTextEditor.document.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sbModified.show();
		sbModified.text = '$(sync) Checking file date';
		sbStatus.hide();
		sbStatus.text = '$(alert) Checked out';
		sp.checkFileState(file).then((data:any) => {
			var modified:Date = new Date(data.TimeLastModified);
			var status:number = data.CheckOutType;
			if (!data.CheckOutType) sbStatus.show();
			sbModified.text = modified <= data.LocalModified ? '$(check) Fresh' : '$(alert) Old';
		});
	});
	// Sync file
	var sync = Commands.registerCommand('sp.sync', () => {
		
	});
	// Upload file
	var upload = Commands.registerCommand('sp.upload', () => {
		
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
		Window.showWarningMessage('You are about to remove all your saved credentials.', 'Cancel', 'Continue').then((selection) => {
			if (selection === 'Continue')
				context.globalState.update('credentials', []).then(() => {
					Window.showInformationMessage('Credentials cache has been reset');
				});
		});
	});
	// Check file sync
	vscode.workspace.onDidOpenTextDocument((e) => {
		vscode.commands.executeCommand('sp.date');
	}, this, context.subscriptions);
	
	context.subscriptions.push(connect, date, sync, upload, checkIn, checkOut, discard, refresh, resetCredentials);
}