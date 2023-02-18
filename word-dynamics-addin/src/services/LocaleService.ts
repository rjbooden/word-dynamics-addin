/* eslint-disable prettier/prettier */

export interface StringLocales {
	newSnippetName: string;
	createSnippet: string;
	selectContent: any;
	close: any;
	selectionDialogTitle: string;
	useSnippets: any;
	useSnippet: any;
	unableDeletingSnippet: string;
	unableStoringSnippet: string;
	unableLoadingSnippets: string;
	results: string;
	searchFor: string;
	search: string;
	searchNotCompleted: string;
	failedSettingContentControlValues: string;
	failedRetrievingIncludedEntity: string;
	failedAddingContentControls: string;
	clickFieldToAdd: string;
	pleaseSideload: string;
	unableToLoadSettings: string;
	title: string;
}

// default locale en-us, also update locale file as a reference when adding a translation
export const strings: StringLocales = {
	title: "Word Dynamics Add-In",
	unableToLoadSettings: "Unable to load add-in settings.",
	pleaseSideload: "Please sideload your add-in to see the app body.",
	clickFieldToAdd: "Click field to add at cursor location:",
	failedAddingContentControls: "Failed adding content controls.",
	failedRetrievingIncludedEntity: "Failed retrieving included entity data.",
	failedSettingContentControlValues: "Failed setting content control values.",
	searchNotCompleted: "Search could not be completed.",
	search: "Search",
	searchFor: "Search for",
	results: "Results:",
	unableLoadingSnippets: "Unable to load snippets",
	unableStoringSnippet: "Unable to store your snippet",
	unableDeletingSnippet: "Unable to delete snippet",
	useSnippet: "Use snippet",
	useSnippets: "Use snippets",
	selectionDialogTitle: "Missing Selection",
	close: "Close",
	selectContent: "Please select content in the Word document.",
	createSnippet: "Create Snippet",
	newSnippetName: "New snippet name:"
};

export class LocaleService {
	public static async getLocale(locale: string): Promise<void> {
		return new Promise<void>((resolve) => {
			// eslint-disable-next-line no-undef
			fetch(`/locales/${locale.toLowerCase()}.json`).then((response) => {
				if (response.status == 200) {
					response.json().then((jsonObject) => {
						for (var property in strings) {
							let locValue = jsonObject[property] as string;
							if (locValue) {
								strings[property] = locValue;
							}
							else {
								// eslint-disable-next-line no-undef
								console.log(`Locale '${locale}' is missing property: ${property}`);
							}
						}
						resolve();
					}).catch(() => {
						// eslint-disable-next-line no-undef
						console.log(`Locale '${locale}' not valid json, using default.`);
						resolve();
					});
				}
				else {
					if (response.status == 404) { // if not found load default locale
						// eslint-disable-next-line no-undef
						console.log(`Locale '${locale}' not found, using default.`);
						resolve();
					}
					else {
						// eslint-disable-next-line no-undef
						console.log(`${response.status} : ${response.statusText}`);
						// eslint-disable-next-line no-undef
						console.log(`Locale '${locale}' not loaded, using default.`);
						resolve();
					}
				}
			}).catch(() => {
				// eslint-disable-next-line no-undef
				console.log(`Locale '${locale}' not loaded, using default.`);
				resolve();
			});
		});
	}
}