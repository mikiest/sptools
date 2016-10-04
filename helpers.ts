import * as vscode from 'vscode';
import Window = vscode.window;

var ctx:vscode.ExtensionContext;
var sbUser:vscode.StatusBarItem = Window.createStatusBarItem(vscode.StatusBarAlignment.Right, 99);
var premCred:helpers.spCredentials;

module helpers {
	export var getContext = (context:vscode.ExtensionContext) => {
		ctx = context;
	}
	export var setCurrentUser = (user:any) => {
		helpers.currentUser.displayName = user.DisplayName;
		helpers.currentUser.email = user.Email;
		sbUser.text = '$(hubot) ' + user.DisplayName;
		sbUser.tooltip = 'SPTools: Authenticated as ' + user.DisplayName;
		sbUser.show();
	};
	export var currentUser = {
		displayName: '',
		email: ''
	}
	export interface spCredentials {
		username: string;
		password: string;
		site?: string;
	}
	// Credentials helper
	export class Credentials {
		stored: Array<helpers.spCredentials>;
		constructor (){
			this.stored = ctx.globalState.get('sp.credentials', []);
		}
		// Prompt for credentials
		private prompt = (site:string) => {
			var credentials = <helpers.spCredentials>{};
			var options:vscode.InputBoxOptions = {
				prompt: 'Username?',
				placeHolder: 'username@domain.com'
			};
			var self = this;
			var promise = new Promise((resolve, reject) => {
				var prompt = () => {
					Window.showInputBox(options).then((selection) => {
						credentials.username = selection;
						options.prompt = 'Password?';
						options.placeHolder = 'Password';
						options.password = true;
						if (!selection) {
							Window.showWarningMessage('Username required.', 'Retry').then((selection) => {
								if (selection === 'Retry') prompt();
							});
							return false;
						}
						Window.showInputBox(options).then((selection) => {
							credentials.password = selection;
							credentials.site = site;
							if (!selection) {
								Window.showWarningMessage('Password required.', 'Retry').then((selection) => {
									if (selection === 'Retry') prompt();
								});
								return false;
							}
							self.store(credentials);
							resolve(credentials);
						});
					});
				};
				prompt();
			});
			return promise;
		}
		// Get credential suggestion using matching URL
		private suggest = (site:string) => {
			// Use memento
			var suggestions:Array<helpers.spCredentials> = this.stored.filter((cred) => {
				return cred.site === site;
			});
			return suggestions;
		}
		// Store credentials in cache
		private store = (credentials:helpers.spCredentials) => {
			if (!vscode.workspace.getConfiguration('sptools').get('storeCredentials')) return false;
			premCred = credentials;
			var exists:boolean = this.stored.filter((cred) => {
				return cred.username === credentials.username;
			}).length > 0;
        	if (exists) return false;
			this.stored.push(credentials);
			ctx.globalState.update('sp.credentials', this.stored);
		}
		// Resolve stored credentials or ask for them
		public get = (site:string) => {
			var suggestions = this.suggest(site);
			var picks = suggestions.map((item) => {
				return item.username || '';
			});
			var credentials:helpers.spCredentials;
			var self = this;
			picks.push('Add credentials');
			var promise = new Promise((resolve, reject) => {
				if (premCred) {
					resolve(premCred);
					return false;
				}
				if (!suggestions.length) {
					self.prompt(site).then((creds) => {
						resolve(creds);
					});
					return false;
				}
				var options = <vscode.QuickPickOptions> {};
				options.placeHolder = 'Authentication to ' + site + ' needed.';
				Window.showQuickPick(picks, options).then((selection) => {
					if(selection === 'Add credentials') {
						self.prompt(site).then((creds) => {
							resolve(creds);
						});
						return false;
					}
					if (!selection) {
						Window.showWarningMessage('Authentication to ' + site + ' needed.', 'Retry').then((selection) => {
							if (selection === 'Retry') self.get(site);
						});
						return false;
					}
					credentials = suggestions.filter((item) => {
						return item.username === selection;
					})[0];
					resolve(credentials);
				});
			});
			return promise;
		}
	}
};

export = helpers;