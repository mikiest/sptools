import * as vscode from 'vscode';
import Window = vscode.window;
import sp = require('./spauth');


export function activate(context: vscode.ExtensionContext) {

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
				var promise = new Promise((resolve, reject) => {
					var promptCredentials = () => {
						options.prompt = 'Username?';
						options.placeHolder = 'username@domain.com';
						Window.showInputBox(options).then((selection) => {
							project.user = selection;
							options.prompt = 'Password?';
							options.placeHolder = 'Password';
							options.password = true;
							Window.showInputBox(options).then((selection) => {
								project.pwd = selection;
								resolve();
							});
						});
					};
					var suggestions = sp.suggestCredentials(project.url);
					var picks = suggestions.map((item) => {
						return item.user;
					});
					picks.push('Add credentials');
					if (suggestions.length) {
						Window.showQuickPick(picks).then((selection) => {
							if(selection === 'Add credentials') {
								promptCredentials();
								return false;
							}
							project.user = selection;
							project.pwd = suggestions.filter((item) => {
								return item.user === selection;
							})[0].password;
							resolve();
						});
					} else {
						promptCredentials();
					}
				});
				promise.then(() => {
					sp.open(project);
				});
			});
		});
	});
	var sbItem:vscode.StatusBarItem;
	// Check file sync
	vscode.workspace.onDidOpenTextDocument((e) => {
		var fileName = e.fileName.split(vscode.workspace.rootPath)[1].split('\\').join('/');
		console.log('Checking: ' + fileName);
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