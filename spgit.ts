var git = require('simple-git');
var emit = require('git-emit');
import * as fs from 'fs';
import * as https from 'https';
import * as constants from 'constants';

var hookFiles = [
	'.git-emit.port',
	'post-rewrite',
	'post-applypatch',
	'post-commit',
	'post-receive',
	'post-checkout',
	'pre-receive',
	'post-merge',
	'pre-auto-gc',
	'applypatch-msg',
	'commit-msg',
	'post-update',
	'pre-applypatch',
	'pre-commit',
	'prepare-commit-msg',
	'pre-push',
	'pre-rebase',
	'update'
];

module spgit {
	export var init = (path:string, callback:() => void) => {
		git(path)
			.init()
			// .add('./*')
			// .commit('Initialized by SP Git')
			// .push('origin', 'master')
			.then(() => {
				var hooks = path + '\\.git\\hooks\\';
				hookFiles.forEach((file) => {
					fs.writeFileSync(hooks + file, '');
				});
				emit(path + '\\.git', (err) => {
					if(err) console.log(err);
				})
				.on('pre-commit', (update) => {
					console.log(update.arguments);
					update.reject();
				});
				callback();
			});
	}
}

export = spgit;