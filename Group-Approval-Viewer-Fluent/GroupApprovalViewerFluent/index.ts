import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Approvals } from "./components/Approvals"

export class GroupApprovalViewerFluent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _notifyOutputChanged: () => void;

	constructor() {}

	public init(
		context: ComponentFramework.Context<IInputs>, 
		notifyOutputChanged: () => void, 
		state: ComponentFramework.Dictionary, 
		container:HTMLDivElement)
	{
		// ### initialization
		this._context = context;
		this._container = container;
		this._notifyOutputChanged = notifyOutputChanged;

		// ### setup dom
		ReactDOM.render(React.createElement(Approvals, this._context), this._container);
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void
	{

	}

	public getOutputs(): IOutputs
	{
		return {};
	}

	public destroy(): void
	{
		ReactDOM.unmountComponentAtNode(this._container);
	}
}