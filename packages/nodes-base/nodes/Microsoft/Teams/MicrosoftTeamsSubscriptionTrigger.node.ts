import {
	createHmac,
	randomUUID,
} from 'crypto';

import {
	parse as urlParse,
} from 'url';

import {
	IHookFunctions,
	IPollFunctions,
	IWebhookFunctions,
} from 'n8n-core';

import {
	IDataObject,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	IWebhookResponseData,
	NodeApiError,
} from 'n8n-workflow';

import {
	decryptAESKey,
	decryptMessage,
	generateMSCert,
	microsoftApiRequest,
	microsoftApiRequestAllItems,
} from './GenericFunctions';
export class MicrosoftTeamsTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'MicrosoftTeams Trigger',
		name: 'microsoftTeamsTrigger',
		icon: 'file:teams.svg',
		group: ['trigger'],
		version: 1,
		description: 'Handle Microsoft Teams events via webhooks',
		defaults: {
			name: 'MicrosoftTeams Trigger',
			color: '#555cc7',
		},
		polling: true,
		inputs: [],
		outputs: ['main'],
		credentials: [
			{
				name: 'microsoftTeamsOAuth2Api',

				required: true,
			},
		],
		webhooks: [
			// {
			// 	name: 'default',
			// 	httpMethod: 'GET',
			// 	responseMode: 'onReceived',
			// 	path: 'teamsData',
			// },
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: 'onReceived',
				path: 'teamsData',
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				options: [
					{
						name: 'Teams',
						value: 'team',
					},
					{
						name: 'Chats',
						value: 'chat',
					},
				],
				default: 'team',
				description: 'The resource to operate on.',
			},
			{
				displayName: 'Return All',
				name: 'returnAll',
				type: 'boolean',
				default: false,
				displayOptions: {
					show: {
						resource: [
							'team',
							'chat',
						],
					},
				},
				description: 'Get messages in all channels in all teams',
			},
			{
				displayName: 'Team ID',
				name: 'teamId',
				required: false,
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getTeams',
				},
				displayOptions: {
					show: {
						resource: [
							'team',
						],
					},
					hide: {
						returnAll: [
							true,
						],
					},
				},
				default: '',
				description: 'Team ID',
			},
			{
				displayName: 'Channel ID',
				name: 'channelId',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getChannels',
					loadOptionsDependsOn: [
						'teamId',
					],
				},
				displayOptions: {
					show: {
						resource: [
							'team',
						],
					},
					hide: {
						returnAll: [
							true,
						],
					},
				},
				default: '',
				description: 'channel ID',
			},
			{
				displayName: 'Chat ID',
				name: 'chatId',
				required: false,
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getChats',
				},
				displayOptions: {
					show: {
						resource: [
							'chat',
						],
					},
					hide: {
						returnAll: [
							true,
						],
					},
				},
				default: '',
				description: 'Team ID',
			},
		],
	};
	methods = {
		loadOptions: {
			// Get all the team's channels to display them to user so that he can
			// select them easily
			async getChannels(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const returnData: INodePropertyOptions[] = [];
				const teamId = this.getCurrentNodeParameter('teamId') as string;
				const { value } = await microsoftApiRequest.call(this, 'GET', `/v1.0/teams/${teamId}/channels`);
				for (const channel of value) {
					const channelName = channel.displayName;
					const channelId = channel.id;
					returnData.push({
						name: channelName,
						value: channelId,
					});
				}
				return returnData;
			},
			// Get all the teams to display them to user so that he can
			// select them easily
			async getTeams(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const returnData: INodePropertyOptions[] = [];
				const { value } = await microsoftApiRequest.call(this, 'GET', '/v1.0/me/joinedTeams');
				for (const team of value) {
					const teamName = team.displayName;
					const teamId = team.id;
					returnData.push({
						name: teamName,
						value: teamId,
					});
				}
				return returnData;
			},
			// Get all the chats to display them to user so that they can
			// select them easily
			async getChats(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const returnData: INodePropertyOptions[] = [];
				const qs: IDataObject = {
					$expand: 'members',
				};
				const { value } = await microsoftApiRequest.call(this, 'GET', '/v1.0/chats', {}, qs);
				const chats = value
								.filter((a: IDataObject) => a.createdDateTime)
								.sort((a: IDataObject, b: IDataObject) => new Date(a.lastUpdatedDateTime as string) > new Date(b.lastUpdatedDateTime as string) ? 1 : -1);
				console.log(`chats: ${JSON.stringify(chats, null, 2)}`);
				for (const chat of value) {
					if (!chat.topic) {
						chat.topic = chat.members
										.filter((member: IDataObject) => member.displayName)
										.map((member: IDataObject) => member.displayName).join(', ');
					}
					const chatName = `${chat.topic || '(no title) - ' + chat.id} (${chat.chatType})`;
					const chatId = chat.id;
					console.log(chatName);
					returnData.push({
						name: chatName,
						value: chatId,
					});
				}
				return returnData;
			},
		},
	};
	// @ts-ignore
	webhookMethods = {
		default: {
			async checkExists(this: IHookFunctions): Promise<boolean> {
				console.log('\n\nCHECK');
				const webhookData = this.getWorkflowStaticData('node');
				const webhookUrl = this.getNodeWebhookUrl('default') as string;
	
				// console.log(`webhookData: ${JSON.stringify(webhookData, null, 2)}`);
				console.log(`webhookUrl: ${JSON.stringify(webhookData, null, 2)}`);

				if (webhookData.webhookId === undefined) {
					console.log('webhookData.webhookId is undefined');
					const responseData = await microsoftApiRequest.call(this, 'GET', `/v1.0/subscriptions`);
					console.log(`responseData: ${JSON.stringify(responseData, null, 2)}`);
					return false;
				}
				try {
					console.log(`Grabbing subscription with id: ${webhookData.webhookId}`);
					const responseData = await microsoftApiRequest.call(this, 'GET', `/v1.0/subscriptions/${webhookData.webhookId}`);
					console.log(`responseData: ${JSON.stringify(responseData, null, 2)}`);
				} catch (error) {
					console.log(`ERROR: grabbing subscription with id: ${webhookData.webhookId}: ${error}`);
					return false;
				}
				// Check if the subscription matches the current webhook
				console.log(`Checking subscription params...`);
				return true;
			},
			async create(this: IHookFunctions): Promise<boolean> {
				console.log('\n\nCREATE');
				const resource = this.getNodeParameter('resource') as string;
				const returnAll = this.getNodeParameter('returnAll') as boolean;
				const webhookUrl = this.getNodeWebhookUrl('default') as string;
				let webhookData = this.getWorkflowStaticData('node');
				const urlParts = urlParse(webhookUrl);

				console.log(`urlParts: ${JSON.stringify(urlParts, null, 2)}`);
				console.log(`webhookData: ${JSON.stringify(webhookData.body, null, 2)}`);
				console.log(`webhookUrl: ${webhookUrl}`);

				let resourceString;
				// const clientState = randomUUID();
				console.log(`Generating CERTIFICATE...`);
				const { fingerprint, cert, keys } = await generateMSCert();
				console.log(`DONE`);
				// console.log(`Generate RSA key pair: ${keys.publicKey}`);
				// const fs = require('fs');
				// fs.writeFile('/tmp/rsa.pub', keys.publicKey, function (err: any) {
				// 	if (err) return console.log(err);
				// });
				// fs.writeFile('/tmp/rsa.key', keys.privateKey, function (err: any) {
				// 	if (err) return console.log(err);
				// });
				// fs.writeFile('/tmp/rsa.crt', cert, function (err: any) {
				// 	if (err) return console.log(err);
				// });

				if (resource === 'team') {
					if (returnAll) {
						resourceString = '/teams/getAllMessages';
					} else {
						const teamId = this.getNodeParameter('teamId') as string;
						const channelId = this.getNodeParameter('channelId') as string;
						resourceString = `/teams/${teamId}/channels/${channelId}/messages`;
					}
				} else if (resource === 'chat') {
					if (returnAll) {
						resourceString = '/chats/getAllMessages';
					} else {
						const chatId = this.getNodeParameter('chatId') as string;
						resourceString = `/chats/${chatId}/messages`;
					}
				}
				// const urlParts = urlParse(webhookUrl);
				const body: IDataObject = {
					changeType: 'created',
					notificationUrl: webhookUrl,
					// notificationUrl: 'https://webhook.site/3bdb5105-7bf6-4321-b9ac-487974c17f48',
					resource: resourceString,
					includeResourceData: true,
					encryptionCertificate: Buffer.from(cert).toString('base64'),
					encryptionCertificateId: fingerprint,
					expirationDateTime:  new Date(Date.now() + 1000 * 60 * 55), // Add 55 minutes	
					clientState: fingerprint,
					// latestSupportedTlsVersion: 'v1_2',
				};

				console.log(`body: ${JSON.stringify(body, null, 2)}`);
				let subscription;
				try {
					subscription = await microsoftApiRequest.call(this, 'POST', `/beta/subscriptions`, body);
					console.log(`subscription: ${JSON.stringify(subscription, null, 2)}`);
				} catch (error) {
					console.log(`error: ${JSON.stringify(error, null, 2)}`);
					// return false;
				}
				
				console.log('Setting webhookdata');
				webhookData = {
					webhookId: subscription?.id,
					fingerprint,
					body,
					cert,
					keys,
				};
				console.log(`\n\nPOST CREATEwebhookData: ${JSON.stringify(webhookData.body, null, 2)}\n\n`);
				return true;
			},
			async delete(this: IHookFunctions): Promise<boolean> {
				console.log('\n\nDELETE');
				let webhookData = this.getWorkflowStaticData('node');

				if (!webhookData) {
					console.log('webhookData is undefined');
					return true;
				}
				try {
					if (webhookData.webhookId) {
						console.log(`Delete subscription with id: ${webhookData.webhookId}`);
						const responseData = await microsoftApiRequest.call(this, 'DELETE', `/v1.0/subscriptions/${webhookData.webhookId}`);
						console.log(`responseData: ${JSON.stringify(responseData, null, 2)}`);
					}
				} catch (error) {
					console.log(`ERROR: delete subscription with id: ${webhookData.webhookId}: ${error}`);
					return false;
				}
			 	webhookData = {};
				return true;
			},
		},
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {

		// Respond with 200 and copy of verification URL param in body for GET
		// Respond with 202 for POST

		const returnData = [];
		const webhookData = this.getWorkflowStaticData('node') as IDataObject;

		const req = this.getRequestObject();
		console.log(`req: ${req}`);

		const { value } = req.body;
		if (value) {

			// Validate JWTs
			// https://docs.microsoft.com/en-us/graph/webhooks-with-resource-data#managing-encryption-keys

			const cert = webhookData.cert;

			// We have resource data, need to loop through and decrypt
			for (const change of value) {

				const fingerprint = change.encryptionCertificateId as string;

				// // Check which key the message is encrypted with
				// if (!webhookData.cert) {
				// 	throw new NodeApiError(this.getNode(), { message: `Cannot find cert with ID ${fingerprint}, we have ${webhookData.fingerprint}` });
				// }

				if (!webhookData.privateKey) {
					throw new NodeApiError(this.getNode(), { message: `Cannot find RSA private key for cert ${fingerprint}` });
				}

				// Decrypt the AES key (dataKey)
				const aesKey = await decryptAESKey(change.encryptionCertificate, webhookData.privateKey);
				// Check HMAC before decrypting
				const hmac = createHmac('sha256', change.content.dataKey);
				if (change.encryptedContent.dataSignature !== hmac.update(change.encryptedContent.data).digest('hex')) {
					throw new NodeApiError(this.getNode(), { message: `HMAC of message incorrect` });
				}
				// Decrypt the message
				const message = await decryptMessage(change.encryptedContent.data, aesKey);
				returnData.push(
					{
						...change,
						...JSON.parse(message),
					});
			}
		} else {
			// Produce normal event
			// OR
			// Request the resource data e.g. the message
			returnData.push(req.body);
		}

		return {
			workflowData: [
				this.helpers.returnJsonArray(returnData),
			],
		};
	}


	async poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
		// REFRESH the subscriptions
		const pollTimes = this.getNodeParameter('pollTimes.item', []) as IDataObject[];
		// const triggerOn = this.getNodeParameter('triggerOn', '') as string;
		// const calendarId = this.getNodeParameter('calendarId') as string;
		const webhookData = this.getWorkflowStaticData('node');
		// const matchTerm = this.getNodeParameter('options.matchTerm', '') as string;

		// Get subscriptions

		const now = new Date(Date.now());

		const startDate = webhookData.lastTimeChecked as string || now;

		const endDate = now;

		console.log(`TEAMS POLLING: ${JSON.stringify(webhookData, null, 2)}`);
		console.log(`POLL TIMES: ${JSON.stringify(pollTimes, null, 2)}`);
		// const qs: IDataObject = {
		// 	showDeleted: false,
		// };

		// if (matchTerm !== '') {
		// 	qs.q = matchTerm;
		// }

		let events;

		// if (triggerOn === 'eventCreated' || triggerOn === 'eventUpdated') {
		// 	Object.assign(qs, {
		// 		updatedMin: startDate,
		// 		orderBy: 'updated',
		// 	});
		// } else if (triggerOn === 'eventStarted' || triggerOn === 'eventEnded') {
		// 	Object.assign(qs, {
		// 		singleEvents: true,
		// 		timeMin: moment(startDate).startOf('second').utc().format(),
		// 		timeMax: moment(endDate).endOf('second').utc().format(),
		// 		orderBy: 'startTime',
		// 	});
		// }

		if (this.getMode() === 'manual') {
			console.log(`MANUAL EXECUTION`);
			events = [{message: 'manual'}];
			// delete qs.updatedMin;
			// delete qs.timeMin;
			// delete qs.timeMax;

			// qs.maxResults = 1;
			// events = await googleApiRequest.call(this, 'GET', `/calendar/v3/calendars/${calendarId}/events`, {}, qs);
			// events = events.items;
		} else {
			console.log(`NORMAL EXECUTION`);
			events = [{message: 'normal'}];
			// events = await googleApiRequestAllItems.call(this, 'items', 'GET', `/calendar/v3/calendars/${calendarId}/events`, {}, qs);
			// if (triggerOn === 'eventCreated') {
			// 	events = events.filter((event: { created: string }) => moment(event.created).isBetween(startDate, endDate));
			// } else if (triggerOn === 'eventUpdated') {
			// 	events = events.filter((event: { created: string, updated: string }) => !moment(moment(event.created).format('YYYY-MM-DDTHH:mm:ss')).isSame(moment(event.updated).format('YYYY-MM-DDTHH:mm:ss')));
			// } else if (triggerOn === 'eventStarted') {
			// 	events = events.filter((event: { start: { dateTime: string } }) => moment(event.start.dateTime).isBetween(startDate, endDate, null, '[]'));
			// } else if (triggerOn === 'eventEnded') {
			// 	events = events.filter((event: { end: { dateTime: string } }) => moment(event.end.dateTime).isBetween(startDate, endDate, null, '[]'));
			// }
		}

		// delete webhookData.certs
		// delete webhookData.body
		webhookData.lastTimeChecked = endDate;

		// if (Array.isArray(events) && events.length) {
		// 	return [this.helpers.returnJsonArray(events)];
		// }

		// if (this.getMode() === 'manual') {
		// 	throw new NodeApiError(this.getNode(), { message: 'No data with the current filter could be found' });
		// }

		return null;
	}
}
