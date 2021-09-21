/* eslint-disable import/no-cycle */
import { IDataObject, IRun } from 'n8n-workflow';
import * as telemetryHelpers from './telemetry/helpers';
import { IInternalHooksClass, IWorkflowBase } from '.';
import { Telemetry } from './telemetry';
import { WorkflowEntity } from './databases/entities/WorkflowEntity';

export class InternalHooksClass implements IInternalHooksClass {
	constructor(private telemetry: Telemetry) {}

	async onServerStarted(n8nVersion: string): Promise<void> {
		await this.telemetry.identify({ n8n_version: n8nVersion });
		await this.telemetry.track('Instance started');
	}

	async onWorkflowSave(workflow: WorkflowEntity): Promise<void> {
		await this.telemetry.track('User saved workflow', {
			workflow_id: workflow.id?.toString(),
			nodes_graph: telemetryHelpers.generateNodesGraph(workflow),
		});
	}

	async onWorkflowActivated(workflow: WorkflowEntity): Promise<void> {
		await this.telemetry.track('User set workflow active status', {
			workflow_id: workflow.id.toString(),
			is_active: workflow.active,
		});
	}

	async onWorkflowTagsUpdated(workflowId: string, tagsCount: number): Promise<void> {
		await this.telemetry.track('User edited workflow tags', {
			workflow_id: workflowId,
			new_tag_count: tagsCount,
		});
	}

	async onWorkflowDeleted(workflowId: string): Promise<void> {
		await this.telemetry.track('User deleted workflow', {
			workflow_id: workflowId,
		});
	}

	async onWorkflowPostExecute(workflow: IWorkflowBase, runData?: IRun): Promise<void> {
		const properties: IDataObject = {
			workflow_id: workflow.id,
		};

		if (runData !== undefined) {
			properties.execution_mode = runData.mode;
			if (runData.mode === 'manual') {
				properties.nodes_graph = telemetryHelpers.generateNodesGraph(workflow);
			}

			properties.success = !!runData.finished;

			if (!properties.success && runData?.data.resultData.error) {
				properties.error_message = runData?.data.resultData.error.message;
				properties.error_node_type = runData?.data.resultData.error.node?.type;

				if (runData.data.resultData.lastNodeExecuted) {
					const lastNode = telemetryHelpers.getNodeTypeForName(
						workflow,
						runData.data.resultData.lastNodeExecuted,
					);
					if (lastNode !== undefined) {
						properties.error_node_type = lastNode.type;
					}
				}
			}
		}

		await this.telemetry.workflowExecutionTrack(properties);
	}
}
