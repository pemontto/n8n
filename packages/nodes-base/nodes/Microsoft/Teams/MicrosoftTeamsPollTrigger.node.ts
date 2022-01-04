import {
	parse as urlParse,
} from 'url';

import {
	IPollFunctions,
} from 'n8n-core';

import {
	IDataObject,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	NodeApiError,
} from 'n8n-workflow';

import {
	microsoftApiRequest,
} from './GenericFunctions';

import * as querystring from 'querystring';

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
					// {
					// 	name: 'Chats',
					// 	value: 'chat',
					// },
				],
				default: 'team',
				description: 'The resource to operate on.',
			},
			// {
			// 	displayName: 'Return All',
			// 	name: 'returnAll',
			// 	type: 'boolean',
			// 	default: false,
			// 	displayOptions: {
			// 		show: {
			// 			resource: [
			// 				'team',
			// 				// 'chat',
			// 			],
			// 		},
			// 	},
			// 	description: 'Get messages in all channels in all teams',
			// },
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
					// hide: {
					// 	returnAll: [
					// 		true,
					// 	],
					// },
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
					// hide: {
					// 	returnAll: [
					// 		true,
					// 	],
					// },
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
					// hide: {
					// 	returnAll: [
					// 		true,
					// 	],
					// },
				},
				default: '',
				description: 'Chat ID',
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
				for (const chat of value) {
					if (!chat.topic) {
						chat.topic = chat.members
										.filter((member: IDataObject) => member.displayName)
										.map((member: IDataObject) => member.displayName).join(', ');
					}
					const chatName = `${chat.topic || '(no title) - ' + chat.id} (${chat.chatType})`;
					const chatId = chat.id;
					returnData.push({
						name: chatName,
						value: chatId,
					});
				}
				return returnData;
			},
		},
	};

	async poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
		const pollTimes = this.getNodeParameter('pollTimes.item', []) as IDataObject[];
		const webhookData = this.getWorkflowStaticData('node');
		const resource = this.getNodeParameter('resource') as string;
		const returnAll = this.getNodeParameter('returnAll', false) as boolean;
		const returnData: IDataObject[] = [];

		const now = new Date(Date.now()).toISOString();
		const startDate = webhookData.lastTimeChecked || now;
		// const startDate = '2021-05-20T20:13:46.153Z';
		const endDate = now;

		const apiVersion = 'v1.0';
		let responseData;
		let resourceString;
		let qs: IDataObject = {};

		// // If we could use a tear down function we could use delta tokens effectively
		// if (webhookData.deltatoken && this.getMode() !== 'manual') {
		// 	console.log(`Got delta token, using that`);
		// 	resourceString = urlParse(webhookData.deltatoken as string).path;
		// } else {
		// 	console.log(`NO delta token, using lastModifiedDateTime`);
		qs.$filter = `lastModifiedDateTime gt ${startDate}`;
		qs.$top = 50;
		if (resource === 'team') {
			if (returnAll) {
				// Will not work as delegated user &  has licensing and payment requirements
				// https://docs.microsoft.com/en-us/graph/api/channel-getallmessages?view=graph-rest-1.0&tabs=http
				resourceString = '/teams/getAllMessages';
			} else {
				const teamId = this.getNodeParameter('teamId') as string;
				const channelId = this.getNodeParameter('channelId') as string;
				resourceString = `/teams/${teamId}/channels/${channelId}/messages/delta`;
			}
		} else if (resource === 'chat') {
			if (returnAll) {
				// Ideally we request the userId and store it once, but that's not possible
				// if we can't remove staticData on trigger activation/deactivation
				if (!webhookData.userId) {
					responseData = await microsoftApiRequest.call(this, 'GET', '/v1.0/me');
					// console.log(`USER IS: ${JSON.stringify(responseData, null, 2)}`);
					// webhookData.userId = responseData.id;
				}

				// Will not work as delegated user &  has licensing and payment requirements
				// https://docs.microsoft.com/en-us/graph/api/chats-getallmessages?view=graph-rest-1.0&tabs=http
				resourceString = `/users/${webhookData.userId}/chats/getAllMessages`;
			} else {
				// Doesn't currently support $filter query param in v1.0 or beta
				const chatId = this.getNodeParameter('chatId') as string;
				resourceString = `/chats/${chatId}/messages`;
			}
		}
		resourceString = `/${apiVersion}${resourceString}`;
		
		if (this.getMode() === 'manual') {
			qs.$top = 1;
			delete qs.$filter;
		}	

		do {
			try {
				responseData = await microsoftApiRequest.call(this, 'GET', resourceString as string, {}, qs);
				returnData.push.apply(returnData, responseData.value);
				if (responseData && '@odata.nextLink' in responseData) {
					delete qs.$filter;
					const resourceString = urlParse(responseData['@odata.nextLink']).path;
					qs = querystring.decode(urlParse(responseData['@odata.nextLink']).query as string);
				}
			} catch (error) {
				throw new NodeApiError(this.getNode(), { message: `${error.message}: ${error.description}` });
			webhookData.deltatoken = responseData['@odata.deltaLink'];
		} while (responseData.value && '@odata.nextLink' in responseData && this.getMode() !== 'manual');

		webhookData.lastTimeChecked = endDate;

		if (returnData.length) {
			return [this.helpers.returnJsonArray(returnData)];
		}

		if (this.getMode() === 'manual') {
			throw new NodeApiError(this.getNode(), { message: 'No data with the current filter could be found' });
		}

		// // The "closeFunction" function gets called by n8n whenever
		// // the workflow gets deactivated and can so clean up.
		// async function closeFunction() {
		// 	console.log(`Calling CLOSE FUNCTION!!!`);
		// 	delete webhookData.lastTimeChecked;
		// 	delete webhookData.userId;
		// 	delete webhookData.deltatoken;
		// }

		return null;
	}
}
