/*
https://github.com/home-assistant/home-assistant-js-websocket/blob/master/lib/socket.ts

*/

import {
	Connection,
	createConnection,
	createLongLivedTokenAuth,
} from 'home-assistant-js-websocket';

import { createSocket } from './socket';

import {
	ICredentialsDecrypted,
	ICredentialTestFunctions,
	IDataObject,
	ILoadOptionsFunctions,
	INodeCredentialTestResult,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	ITriggerFunctions,
	ITriggerResponse,
} from 'n8n-workflow';

import {
	getHomeAssistantEntities,
} from './GenericFunctions';

export class HomeAssistantTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Home Assistant Trigger',
		name: 'homeAssistantTrigger',
		icon: 'file:homeAssistant.svg',
		group: ['trigger'],
		version: 1,
		subtitle: '={{$parameter["event"]}}',
		description: 'Listens to Home Assistant events',
		defaults: {
			name: 'Home Assistant Trigger',
		},
		inputs: [],
		outputs: ['main'],
		credentials: [
			{
				name: 'homeAssistantApi',
				required: true,
				testedBy: 'homeAssistantApiTest',
			},
		],
		properties: [
			{
				displayName: 'Event',
				name: 'eventType',
				type: 'options',
				default: 'subscribe_trigger',
				required: true,
				description: 'The type of the events to listen to',
				options: [
					{
						name: '*',
						value: '*',
						description: 'Any time any event is triggered (Wildcard Event)',
					},
					{
						name: 'Subscribe Trigger',
						value: 'subscribe_trigger',
					},
					{
						name: 'Automation Reloaded',
						value: 'automation_reloaded',
					},
					{
						name: 'Automation Triggered',
						value: 'automation_triggered',
					},
					{
						name: 'Call Service',
						value: 'call_service',
					},
					{
						name: 'Component Loaded',
						value: 'component_loaded',
					},
					{
						name: 'Core Config Updated',
						value: 'core_config_updated',
					},
					{
						name: 'Data Entry Flow Progressed',
						value: 'data_entry_flow_progressed',
					},
					{
						name: 'Homeassistant Close',
						value: 'homeassistant_close',
					},
					{
						name: 'Homeassistant Final Write',
						value: 'homeassistant_final_write',
					},
					{
						name: 'Homeassistant Start',
						value: 'homeassistant_start',
					},
					{
						name: 'Homeassistant Started',
						value: 'homeassistant_started',
					},
					{
						name: 'Homeassistant Stop',
						value: 'homeassistant_stop',
					},
					{
						name: 'Logbook Entry',
						value: 'logbook_entry',
					},
					{
						name: 'Scene Reloaded',
						value: 'scene_reloaded',
					},
					{
						name: 'Script Started',
						value: 'script_started',
					},
					{
						name: 'Service Registered',
						value: 'service_registered',
					},
					{
						name: 'Service Removed',
						value: 'service_removed',
					},
					{
						name: 'State Changed',
						value: 'state_changed',
					},
					{
						name: 'Themes Updated',
						value: 'themes_updated',
					},
					{
						name: 'User Added',
						value: 'user_added',
					},
					{
						name: 'User Removed',
						value: 'user_removed',
					},
				],
			},
			{
				displayName: 'Entity ID',
				name: 'entityId',
				type: 'multiOptions',
				typeOptions: {
					loadOptionsMethod: 'getAllEntities',
				},
				displayOptions: {
					show: {
						eventType: [
							'subscribe_trigger',
						],
					},
				},
				required: false,
				default: [],
			},
			{
				displayName: 'From State',
				name: 'fromState',
				type: 'string',
				displayOptions: {
					show: {
						eventType: [
							'subscribe_trigger',
						],
					},
				},
				required: false,
				default: '',
				description: 'Comma separated list of states transitioned from',
			},
			{
				displayName: 'To State',
				name: 'toState',
				type: 'string',
				displayOptions: {
					show: {
						eventType: [
							'subscribe_trigger',
						],
					},
				},
				required: false,
				default: '',
				description: 'Comma separated list of states transitioned to',
			},
		],
	};

	methods = {
		credentialTest: {
			async homeAssistantApiTest(this: ICredentialTestFunctions, credential: ICredentialsDecrypted): Promise<INodeCredentialTestResult> {
				const credentials = credential.data;
				const options = {
					method: 'GET',
					headers: {
						Authorization: `Bearer ${credentials!.accessToken}`,
					},
					uri: `${credentials!.ssl === true ? 'https' : 'http'}://${credentials!.host}:${ credentials!.port || '8123' }/api/`,
					json: true,
					timeout: 5000,
				};
				try {
					const response = await this.helpers.request(options);
					if (!response.message) {
						return {
							status: 'Error',
							message: `Token is not valid: ${response.error}`,
						};
					}
				} catch (error) {
					return {
						status: 'Error',
						message: `${error.statusCode === 401 ? 'Token is' : 'Settings are'} not valid: ${error}`,
					};
				}

				return {
					status: 'OK',
					message: 'Authentication successful!',
				};
			},
		},
		loadOptions: {
			async getAllEntities(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				return await getHomeAssistantEntities.call(this);
			},
		},
	};

	async trigger(this: ITriggerFunctions): Promise<ITriggerResponse> {

		const eventType = this.getNodeParameter('eventType') as string;

		const credentials = await this.getCredentials('homeAssistantApi');
		const hassUrl = `${credentials.ssl === true ? 'wss' : 'ws'}://${credentials.host}:${credentials.port}/api/websocket`;

		// console.log('Creating Long Lived Home Assistant Token...');
		const auth = createLongLivedTokenAuth(
			hassUrl,
			credentials.accessToken as string,
		);

		// console.log('Connecting to Home Assistant...');
		let conn: Connection;
		try {
			conn = await createConnection({
				auth,
				createSocket: async () => createSocket(auth, hassUrl),
			});
		} catch (error) {
			throw new Error(`Could not connect to Home Assistant: ${error}`);
		}


		const haEvent = (event: IDataObject) => {
			// console.log(JSON.stringify(event));
			this.emit([[{ json: event}]]);
		};

		const haTriggerEvent = (event: IDataObject) => {
			// console.log(JSON.stringify(event));
			this.emit(
				[[{
					json: {
						...(event.variables as IDataObject).trigger as IDataObject,
						context: event.context,
					},
				}]]);
		};
		// console.log(`Connected to to Home Assistant ${JSON.stringify(auth)}`);

		console.log(`Subscribing to ${eventType}...`);
		if (eventType === 'subscribe_trigger') {
			const entityIds = this.getNodeParameter('entityId') as string[];
			const fromState = this.getNodeParameter('fromState') as string;
			const toState = this.getNodeParameter('toState') as string;
			const triggerCondition: IDataObject = {
				platform: 'state',
			};
			if (fromState) {
				triggerCondition.from = fromState.split(',');
			}
			if (toState) {
				triggerCondition.to = toState.split(',');
			}
			// console.log(`Subscribing to ${entityIds}`);
			for (const entityId of entityIds) {
				console.log(`entityId: ${entityId}`);
				triggerCondition.entity_id = entityId;
				console.log(`triggerCondition: ${JSON.stringify(triggerCondition)}`);
				conn.subscribeMessage(haTriggerEvent, {
					type: eventType,
					trigger: triggerCondition,
				});
			}
		} else if ( eventType === '*') {
			// console.log(`Subscribing to all events`);
			conn.subscribeEvents(haEvent);
		} else {
			conn.subscribeEvents(haEvent, eventType);
		}

		// The "closeFunction" function gets called by n8n whenever
		// the workflow gets deactivated and can so clean up.
		async function closeFunction() {
			// console.log('DISCONNECTING from Home Assistant...');
			conn.close();
		}

		return {
			closeFunction,
		};
	}

}
