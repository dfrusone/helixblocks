import { Plugin } from '../../components/plugin/plugin.js';
import { google } from 'googleapis';
import TurndownService from 'turndown';

// Initialize Turndown Service
const turndownService = new TurndownService();

// Google Docs API setup
const SCOPES = ['https://www.googleapis.com/auth/documents.readonly'];
const TOKEN_PATH = 'token.json';

// Function to initialize Google Docs API
async function initializeGoogleDocsApi() {
    const auth = new google.auth.GoogleAuth({
        keyFile: 'path/to/your/credentials.json',
        scopes: SCOPES,
    });
    const authClient = await auth.getClient();
    const docs = google.docs({ version: 'v1', auth: authClient });
    return docs;
}

// Function to fetch document content and convert to Markdown
async function fetchAndConvertToMarkdown(docId) {
    const docs = await initializeGoogleDocsApi();
    const res = await docs.documents.get({ documentId: docId });
    const content = res.data.body.content;
    let markdown = '';
    content.forEach(element => {
        if (element.paragraph) {
            const text = element.paragraph.elements.map(e => e.textRun.content).join('');
            markdown += turndownService.turndown(text) + '\n\n';
        }
    });
    return markdown;
}

/**
 * Creates the preview plugin
 * @param {AppStore} appStore The app store
 * @returns {Plugin} The preview plugin
 */
export function createPreviewPlugin(appStore) {
    return new Plugin({
            id: 'edit-preview',
            condition: (store) => store.isEditor(),
            button: {
                text: appStore.i18n('preview'),
                action: async () => {
                    const { location } = appStore;
                    const status = await appStore.fetchStatus(false, true, true);
                    if (status.edit && status.edit.sourceLocation
                        && status.edit.sourceLocation.startsWith('onedrive:')
                        && !location.pathname.startsWith('/:x:/')) {
                        const mac = navigator.platform.toLowerCase().includes('mac') ? '_mac' : '';
                        appStore.showToast(appStore.i18n(`preview_onedrive${mac}`));
                    } else if (status.edit.sourceLocation?.startsWith('gdrive:')) {
                        const { contentType, sourceLocation } = status.edit;
                        const isGoogleDocMime = contentType === 'application/vnd.google-apps.document';
                        const isGoogleSheetMime = contentType === 'application/vnd.google-apps.spreadsheet';
                        const neitherGdocOrGSheet = !isGoogleDocMime && !isGoogleSheetMime;

                        if (neitherGdocOrGSheet) {
                            const isMsDocMime = contentType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
                            const isMsExcelSheet = contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                            let errorKey = 'error_preview_not_gdoc_generic'; // show generic message by default

                            if (isMsDocMime) {
                                errorKey = 'error_preview_not_gdoc_ms_word';
                            } else if (isMsExcelSheet) {
                                errorKey = 'error_preview_not_gsheet_ms_excel';
                            }

                            appStore.showToast(
                                appStore.i18n(errorKey),
                                'negative',
                                () => appStore.closeToast(),
                            );

                            return;
                        }

                        // Convert Google Doc to Markdown
                        if (isGoogleDocMime) {
                            const docId = sourceLocation.split(':').pop();
                            const markdown = await fetchAndConvertToMarkdown(docId);
                            console.log('Markdown content:', markdown);
                            // Save or handle the markdown as needed
                        }
                    }
                    if (location.pathname.startsWith('/:x:/')) {
                        window.sessionStorage.setItem('aem-sk-preview', JSON.stringify({
                            previewPath: status.webPath,
                            previewTimestamp: Date.now(),
                        }));
                        appStore.reloadPage();
                    } else {
                        appStore.updatePreview();
                    }
                },
                isEnabled: (store) => store.isAuthorized('preview', 'write')
                    && store.status.webPath,
            },
            callback: () => {
                const { previewPath, previewTimestamp } = JSON
                    .parse(window.sessionStorage.getItem('aem-sk-preview') || '{}');
                window.sessionStorage.removeItem('aem-sk-preview');
                if (previewTimestamp < Date.now() + 60000) {
                    const { status } = appStore;
                    if (status.webPath === previewPath && appStore.isAuthorized('preview', 'write')) {
                        appStore.updatePreview();
                    } else {
                        appStore.closeToast();
                    }
                }
            },
        },
        appStore);
}
