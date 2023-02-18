/* eslint-disable prettier/prettier */
import Wait from "../helper/Wait";

export enum ContentSource {
	None = 1,
	Text = 2,
	Object = 3
}

export class TokenService {

	private static accessToken;
	private static accessTokenSet: Date;
	private static retrieving: boolean;

	private static clearAccessToken() {
		let now = new Date();
		// Only clear if the token is older than 10 seconds to prevent 
		// clearing multiple times within this timeframe
		if ((now.getTime() - this.accessTokenSet.getTime()) > 10000) {
			this.accessToken = null;
		}
	}

	private static async ensureToken(): Promise<string> {
		if (!this.accessToken) {
			return new Promise<string>((resolve, reject) => {
				if (this.retrieving) {
					Wait(800, 10, () => { // max wait time 8 seconds
						return !this.accessToken;
					}).then(() => {
						resolve(this.accessToken);
					}).catch((reason) => { reject(reason); });
				}
				else {
					this.retrieving = true;
					let options = {
						allowSignInPrompt: true,
						allowConsentPrompt: true,
						forMSGraphAccess: true,
					};
					// eslint-disable-next-line no-undef
					Office.auth.getAccessToken(options).then((accessToken) => {
						this.accessToken = accessToken;
						this.accessTokenSet = new Date();
						this.retrieving = false;
						resolve(this.accessToken);
					}).catch((reason) => {
						if (typeof reason == 'object' && reason.code) {
							reject(`${reason.code} - ${reason.message}`)
						}
						else {
							reject(reason);
						}
					});
				}
			});
		}
		return Promise.resolve(this.accessToken);
	}

	public static AuthenticatedRequest<T>(input: RequestInfo | URL, init?: RequestInit, fromContent: ContentSource = ContentSource.Object, retryCount: number = 0, retry?: boolean): Promise<T> {
		return new Promise<T>((resolve, reject) => {
			TokenService.ensureToken().then((accessToken) => {
				// eslint-disable-next-line no-undef
				if (init.headers) {
					let newHeaders = {};
					for (const property in init.headers) {
						if (property != 'Authorization') {
							newHeaders[property] = init.headers[property];
						}
					}
					newHeaders['Authorization'] = "bearer " + accessToken;
					init.headers = newHeaders;
				}
				else {
					init.headers = {
						'Authorization': "bearer " + accessToken
					};
				}
				// eslint-disable-next-line no-undef
				fetch(input, init).then((response) => {
					if (response.status >= 200 && response.status < 300) {
						if (fromContent == ContentSource.Object) {
							response.json().then((jsonObject) => {
								resolve(jsonObject as T);
							}).catch((reason) => {
								reject(reason);
							});
						} else if (fromContent == ContentSource.Text) {
							response.text().then((text) => {
								resolve(text as T);
							}).catch((reason) => {
								reject(reason);
							});
						}
						else {
							resolve(null as T);
						}
					}
					else if (response.status == 401) {
						if (!retry) {
							TokenService.clearAccessToken();
							// eslint-disable-next-line no-undef
							setTimeout(() => { // wait a bit before retrying
								this.AuthenticatedRequest(input, init, fromContent, retryCount, true)
									.then((value) => { resolve(value as T); })
									.catch((reason) => { reject(reason); });
							}, 200);
						}
						else {
							response.text().then((text) => {
								reject(`${response.status} : ${text}`);
							}).catch(() => {
								reject(`${response.status}`)
							});
						}
					}
					else if (response.status == 429 || response.status == 503) {
						if (retryCount > 20) {
							if (response.status == 429) {
								reject('Request throttling exceeded limit.');
							}
							else {
								reject('Service unavailable, retry exceeded limit.');
							}
						}
						else {
							let waitSeconds = 1;
							let retryAfter = response.headers.get('Retry-After');
							if (retryAfter) {
								waitSeconds = parseInt(retryAfter);
							}
							// eslint-disable-next-line no-undef
							setTimeout(() => {
								retryCount++;
								this.AuthenticatedRequest(input, init, fromContent, retryCount, false)
									.then((value) => { resolve(value as T); })
									.catch((reason) => { reject(reason); });
							}, waitSeconds * 1000);
						}
					}
					else {
						reject(`${response.status} : ${response.statusText}`);
					}
				}).catch((reason) => { reject(reason); });
			}).catch((reason) => { reject(reason); });
		});
	}

}