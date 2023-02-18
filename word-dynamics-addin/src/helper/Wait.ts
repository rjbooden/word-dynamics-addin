/* eslint-disable prettier/prettier */

export default async function Wait(milliseconds: number = 200, // initial wait time
	maxWaitCount: number = 25,  // max retry count (200 * 25 = 5 seconds)
	needsRetry: () => boolean,
	waitCount: number = 0) {    // wait counter
	return new Promise<void>((resolve, reject) => {
		// eslint-disable-next-line no-undef
		setTimeout(() => {
			if (waitCount > maxWaitCount) {
				reject('Max wait time reached');
			}
			else {
				if (needsRetry()) {
					waitCount++;
					Wait(milliseconds, maxWaitCount, needsRetry, waitCount).then(() => {
						resolve();
					}).catch((reason) => {
						reject(reason);
					});
				}
				else {
					resolve();
				}
			}
		}, milliseconds);
	});
}