const configuration = await window.webkit.messageHandlers.desktop.postMessage({type: 'configuration'});
if (!document.querySelector('.start-screen .open-primary')) throw new Error('Missing file opener');
if (document.querySelector('.editor-stage')) throw new Error('A document opened at launch');
return {format: configuration.format, start: true};
