import * as vscode from 'vscode';
import Window = vscode.window;
import sp = require('./spauth');
import helpers = require('./helpers');


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
	var sbItem:vscode.StatusBarItem;
	// Check file sync
	vscode.workspace.onDidOpenTextDocument((e) => {
		var fileName = e.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		if (!sbItem) {
			sbItem = Window.createStatusBarItem(vscode.StatusBarAlignment.Left);
			sbItem.show();
		}
		sbItem.text = '$(sync) Checking';
		sbItem.show();
		sp.checkFile(fileName).then((uptodate) => {
			sbItem.text = uptodate ? '$(check) Fresh' : '$(alert) Old';
		});
	}, this, context.subscriptions);
	
	context.subscriptions.push(disposable);
}