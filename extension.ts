// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import Window = vscode.window;
import sp = require('./spauth');

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
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
	
	context.subscriptions.push(disposable);
}