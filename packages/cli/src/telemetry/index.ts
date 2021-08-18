/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
import TelemetryClient = require('@rudderstack/rudder-sdk-node');
import { IDataObject } from 'n8n-workflow';
import config = require('../../config');

export class Telemetry {
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	private client?: any;

	private instanceId: string;

	constructor(instanceId: string) {
		this.instanceId = instanceId;

		const enabled = config.get('telemetry.enabled') as boolean;
		if (enabled) {
			this.client = new TelemetryClient(
				config.get('telemetry.config.backend.key') as string,
				config.get('telemetry.config.backend.url') as string,
			);
		}
	}

	async identify(traits?: IDataObject): Promise<void> {
		if (this.client) {
			this.client.identify({
				userId: this.instanceId,
				anonymousId: '000000000000',
				traits,
			});
		}
	}

	async track(eventName: string, properties?: IDataObject): Promise<void> {
		if (this.client) {
			this.client.track({
				userId: this.instanceId,
				event: eventName,
				anonymousId: '000000000000',
				properties,
			});
		}
	}
}
