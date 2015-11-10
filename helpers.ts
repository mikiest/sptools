import * as vscode from 'vscode';
import Window = vscode.window;

var ctx:vscode.ExtensionContext;

module helpers {
	export var getContext = (context:vscode.ExtensionContext) => {
		ctx = context;
	};
	export interface spCredentials {
		username: string;
		password: string;
		site?: string;
	}
	export class Credentials {
		stored: Array<helpers.spCredentials>;
		constructor (){
			this.stored = ctx.globalState.get<Array<helpers.spCredentials>>('credentials');
		}
		private prompt = () => {
			var credentials:helpers.spCredentials;
			var options: vscode.InputBoxOptions = {
				prompt: 'Username?',
				placeHolder: 'username@domain.com'
			};
			var promise = new Promise((resolve, reject) => {
				Window.showInputBox(options).then((selection) => {
					credentials.username = selection;
					options.prompt = 'Password?';
					options.placeHolder = 'Password';
					options.password = true;
					Window.showInputBox(options).then((selection) => {
						credentials.password = selection;
						resolve(credentials);
					});
				});
			});
			return promise;
		}
		private suggest = (site:string) => {
			// Use memento
			var suggestions:Array<helpers.spCredentials> = this.stored.filter((cred) => {
				return cred.site === site;
			});
			return suggestions;
		}
		private store = (credentials:helpers.spCredentials) => {
			var exists:boolean = this.stored.filter((cred) => {
				return cred.username === credentials.username;
			}).length > 0;
        	if (exists) return false;
			this.stored.push(credentials);
			ctx.globalState.update('credentials', this.stored);
		}
		public get = (site:string) => {
			var suggestions = this.suggest(site);
			var picks = suggestions.map((item) => {
				return item.username;
			});
			var credentials:helpers.spCredentials;
			var self = this;
			picks.push('Add credentials');
			var promise = new Promise((resolve, reject) => {
				if (!suggestions.length) {
					self.prompt().then((creds) => {
						resolve(credentials);
					});
					return false;
				}
				Window.showQuickPick(picks).then((selection) => {
					if(selection === 'Add credentials') {
						self.prompt().then((creds) => {
							resolve(credentials);
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