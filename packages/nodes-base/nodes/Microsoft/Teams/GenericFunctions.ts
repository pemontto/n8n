import {
	pki,
} from 'node-forge';

import {
	constants,
	createDecipheriv,
	privateDecrypt,
} from 'crypto';

import {
	OptionsWithUri,
} from 'request';

import {
	IExecuteFunctions,
	IExecuteSingleFunctions,
	IHookFunctions,
	ILoadOptionsFunctions,
} from 'n8n-core';

import {
	IDataObject, IPollFunctions, NodeApiError,
} from 'n8n-workflow';

export async function microsoftApiRequest(this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions | IPollFunctions, method: string, resource: string, body: any = {}, qs: IDataObject = {}, uri?: string, headers: IDataObject = {}): Promise<any> { // tslint:disable-line:no-any
	const options: OptionsWithUri = {
		headers: {
			'Content-Type': 'application/json',
		},
		method,
		body,
		qs,
		uri: uri || `https://graph.microsoft.com${resource}`,
		json: true,
	};
	try {
		if (Object.keys(headers).length !== 0) {
			options.headers = Object.assign({}, options.headers, headers);
		}
		//@ts-ignore
		return await this.helpers.requestOAuth2.call(this, 'microsoftTeamsOAuth2Api', options);
	} catch (error) {
		throw new NodeApiError(this.getNode(), error);
	}
}

export async function generateMSCert(modulusLength = 2048): Promise<{ fingerprint: string, cert: string, keys: any }> {
	const keys = pki.rsa.generateKeyPair({bits: modulusLength, workers: -1});
	const cert = pki.createCertificate();
	cert.publicKey = keys.publicKey;
	cert.validity.notAfter.setFullYear(cert.validity.notBefore.getFullYear() + 1); // adding 1 year of validity from no
	cert.sign(keys.privateKey);
	const fingerprint = pki.getPublicKeyFingerprint(cert.publicKey, {type: 'SubjectPublicKeyInfo', encoding: 'hex'});
	// convert a Forge certificate to PEM
	const pem = pki.certificateToPem(cert);
	return {
		fingerprint,
		cert: pem,
		keys,
	};
	// return generateKeyPairSync('rsa', {
	// 	modulusLength,
	// 	publicKeyEncoding: {
	// 		type: 'spki',
	// 		format: 'pem',
	// 	},
	// 	privateKeyEncoding: {
	// 		type: 'pkcs8',
	// 		format: 'pem',
	// 		cipher,
	// 		passphrase,
	// 	},
	// });
}

export async function decryptAESKey(encryptedKey: string, privateKey: any): Promise<string> {

	return privateKey.decrypt(encryptedKey, 'RSA-OAEP');
	// const decrypted = privateDecrypt({
	// 		key: privateKey,
	// 		padding: constants.RSA_PKCS1_OAEP_PADDING,
	// 		passphrase,
	// 	},
	// 	Buffer.from(encryptedKey, 'base64'),
	// );
	// return decrypted.toString('utf8');
}

export async function decryptMessage(encryptedData: string, key: string, cipher = 'aes-256-cbc'): Promise<string> {
	// Create IV from first 16 bytes of the key
	const iv = key.slice(0, 16);
	const decryptCipher = createDecipheriv(cipher, encryptedData, iv);
	return Buffer.concat([decryptCipher.update(encryptedData, 'base64'), decryptCipher.final()]).toString('utf8');
}

export async function microsoftApiRequestAllItems(this: IExecuteFunctions | ILoadOptionsFunctions, propertyName: string, method: string, endpoint: string, body: any = {}, query: IDataObject = {}): Promise<any> { // tslint:disable-line:no-any

	const returnData: IDataObject[] = [];

	let responseData;
	let uri: string | undefined;

	do {
		responseData = await microsoftApiRequest.call(this, method, endpoint, body, query, uri);
		uri = responseData['@odata.nextLink'];
		returnData.push.apply(returnData, responseData[propertyName]);
		if (query.limit && query.limit <= returnData.length) {
			return returnData;
		}
	} while (
		responseData['@odata.nextLink'] !== undefined
	);

	return returnData;
}

export async function microsoftApiRequestAllItemsSkip(this: IExecuteFunctions | ILoadOptionsFunctions, propertyName: string, method: string, endpoint: string, body: any = {}, query: IDataObject = {}): Promise<any> { // tslint:disable-line:no-any

	const returnData: IDataObject[] = [];

	let responseData;
	query['$top'] = 100;
	query['$skip'] = 0;

	do {
		responseData = await microsoftApiRequest.call(this, method, endpoint, body, query);
		query['$skip'] += query['$top'];
		returnData.push.apply(returnData, responseData[propertyName]);
	} while (
		responseData['value'].length !== 0
	);

	return returnData;
}
