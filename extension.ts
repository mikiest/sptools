import * as vscode from 'vscode';
import Window = vscode.window;
import sp = require('./spcore');
import helpers = require('./helpers');
import * as fs from 'fs';


export function activate(context: vscode.ExtensionContext) {
	
	sp.getContext(context);

	var disposable = vscode.commands.registerCommand('sp.connect', () => {
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
	var sbModified:vscode.StatusBarItem = Window.createStatusBarItem(vscode.StatusBarAlignment.Left);
	var sbStatus:vscode.StatusBarItem = Window.createStatusBarItem(vscode.StatusBarAlignment.Left);
	// Check file sync
	vscode.workspace.onDidOpenTextDocument((e) => {
		var fileName = e.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		sbModified.show();
		sbModified.text = '$(sync) Checking file date';
		sbStatus.hide();
		sbStatus.text = '$(alert) Checked out';
		sp.checkFileState(fileName).then((data:any) => {
			var modified:Date = new Date(data.TimeLastModified);
			var status:number = data.CheckOutType;
			if (!data.CheckOutType) sbStatus.show();
			sbModified.text = modified <= data.LocalModified ? '$(check) Fresh' : '$(alert) Old';
		});
	}, this, context.subscriptions);
	
	context.subscriptions.push(disposable);
}