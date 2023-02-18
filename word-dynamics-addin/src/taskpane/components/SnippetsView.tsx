/* eslint-disable prettier/prettier */
import { Dialog, DialogFooter, DialogType, IconButton, PrimaryButton, Spinner, SpinnerSize, Stack, TextField } from "@fluentui/react";
import * as React from "react";
import { strings } from "../../services/LocaleService";
import { SettingsService, Snippet } from "../../services/SettingsService";
import Wait from "../../helper/Wait";
import ClickableListItem from "./ClickableListItem";

export interface SnippetProps {
	onError: (errorMessage: string, errorInfo?: string) => void;
}

export interface SnippetState {
	newSnippetName: string;
	snippets: Snippet[];
	isLoading: boolean;
	isSaving: boolean;
	hideDialog: boolean;
}

export default class SnippetsView extends React.Component<SnippetProps, SnippetState> {
	private isWorking: boolean = false;

	constructor(props, state) {
		super(props, state);
		this.state = {
			newSnippetName: '',
			snippets: [],
			isLoading: true,
			isSaving: false,
			hideDialog: true,
		}
	}

	componentDidMount(): void {
		SettingsService.loadSnippets().then((snippets) => {
			this.setState({ snippets, isLoading: false });
		}).catch((reason) => {
			// eslint-disable-next-line no-undef
			console.log(reason);
			this.props.onError(strings.unableLoadingSnippets, reason);
			this.setState({ isLoading: false });
		});
	}

	// eslint-disable-next-line no-undef
	ensureFullyIncludedContentControls = async (context: Word.RequestContext, range: Word.Range) => {
		context.load(range, ['contentControls', 'contentControls.items', 'contentControls.items.length']);
		await context.sync();

		let startRange = null;
		let endRange = null;
		const items = range.contentControls.items;
		if (items.length > 0) {
			// eslint-disable-next-line no-undef
			startRange = items[0].getRange(Word.RangeLocation.whole);
			// eslint-disable-next-line office-addins/load-object-before-read
			if (items.length > 1) {
				// eslint-disable-next-line no-undef
				endRange = items[items.length - 1].getRange(Word.RangeLocation.whole);
			}
			await context.sync();
			range = range.expandTo(startRange);
			if (endRange) {
				range = range.expandTo(endRange);
			}
			await context.sync();
		}
		return range;
	}

	onCreateSnippet = () => {
		this.setState({ isSaving: true });
		this.isWorking = true;
		// eslint-disable-next-line no-undef
		Word.run(async (context) => {
			try {
				// Gets the content of the range
				// eslint-disable-next-line no-undef
				let range = context.document.getSelection().getRange();
				context.load(range, 'isEmpty');
				await context.sync();

				// eslint-disable-next-line office-addins/load-object-before-read
				if (!range.isEmpty) {
					range = await this.ensureFullyIncludedContentControls(context, range);
					const ooxml = range.getOoxml();
					await context.sync();
					let snippets: Snippet[] = JSON.parse(JSON.stringify(this.state.snippets));
					// eslint-disable-next-line office-addins/load-object-before-read
					snippets.push({ title: this.state.newSnippetName, content: ooxml.value });
					SettingsService.saveSnippets(snippets).then(() => {
						this.setState({ snippets, isSaving: false, newSnippetName: '' });
						this.props.onError(null);
						this.isWorking = false;
					}).catch((reason) => {
						// eslint-disable-next-line no-undef
						console.log(reason);
						this.props.onError(strings.unableStoringSnippet, reason);
						this.setState({ isSaving: false });
						this.isWorking = false;
					});
				}
				else {
					this.setState({ isSaving: false, hideDialog: false });
					this.isWorking = false;
				}
			} catch (e) {
				// eslint-disable-next-line no-undef
				console.log(e);
				this.props.onError(e);
				this.setState({ isSaving: false });
				this.isWorking = false;
			}
		});
	}

	closeDialog = () => {
		this.setState({ hideDialog: true })
	}

	deleteSnippet = (s: Snippet) => {
		const handleDelete = () => {
			this.isWorking = true;
			let deepCopy: Snippet[] = JSON.parse(JSON.stringify(this.state.snippets));
			let snippets = deepCopy.filter((item) => { return item.title != s.title });
			SettingsService.saveSnippets(snippets).then(() => {
				this.setState({ snippets });
				this.props.onError(null);
				this.isWorking = false;
			}).catch((reason) => {
				this.setState({ snippets }); // to clear deleting spinner
				// eslint-disable-next-line no-undef
				console.log(reason);
				this.props.onError(strings.unableDeletingSnippet, reason);
				this.isWorking = false;
			});
		};
		if (this.isWorking) { // handle multiple deletes sequentially
			Wait(200, 80, () => {
				return this.isWorking;
			}).then(() => {
				handleDelete();
			}).catch((reason) => {
				// eslint-disable-next-line no-undef
				console.log(reason);
				this.props.onError(strings.unableDeletingSnippet, reason);
			});
		}
		else {
			handleDelete();
		}
	}

	insertSnippet = (Snippet: Snippet, notifyLoaded) => {
		// eslint-disable-next-line no-undef
		Word.run(async (context) => {
			try {
				// Get range
				const range = context.document.getSelection().getRange();
				await context.sync();

				// Set the content of the range
				range.insertOoxml(Snippet.content, "Replace");
				await context.sync();
				notifyLoaded();
			} catch (e) {
				// eslint-disable-next-line no-undef
				console.log(e);
				this.props.onError(e);
				notifyLoaded();
			}
		});
	}

	render() {

		let snippetItems = this.state.snippets.map((s, i) => {
			return (<ClickableListItem
				label={s.title}
				iconName="Paste"
				key={s.title + i}
				showLoading={true}
				onDelete={() => { this.deleteSnippet(s); }}
				onClick={(notifyLoaded) => { this.insertSnippet(s, notifyLoaded); }} />);
		});

		const enableCopy = this.state.newSnippetName != ''
			&& (this.state.snippets.length === 0
				|| null == (this.state.snippets.find((s) => {
					return s.title == this.state.newSnippetName
				})));

		let useText = snippetItems.length == 1 ? strings.useSnippet : strings.useSnippets

		const dialogContentProps = {
			type: DialogType.normal,
			title: strings.selectionDialogTitle,
			closeButtonAriaLabel: strings.close,
			subText: strings.selectContent,
		};

		return (
			<main className="ms-welcome__main">
				<h2 className="ms-font-m ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{strings.createSnippet}</h2>
				<Stack tokens={{ childrenGap: 8 }} className="snippet-stack" horizontal verticalAlign="baseline">
					<TextField label={strings.newSnippetName} value={this.state.newSnippetName} className="snippet-name" onChange={(_e, v) => { this.setState({ newSnippetName: v }) }} />
					{
						this.state.isSaving ?
							<Spinner size={SpinnerSize.small} className="saving-spinner" />
							: <IconButton
								iconProps={{ iconName: 'SaveTemplate' }}
								onClick={this.onCreateSnippet}
								disabled={!enableCopy}
								className="snippet-button"
							/>
					}
				</Stack>
				<Dialog
					hidden={this.state.hideDialog}
					onDismiss={this.closeDialog}
					dialogContentProps={dialogContentProps}
					modalProps={{ isBlocking: true }}
				>
					<DialogFooter>
						<PrimaryButton onClick={this.closeDialog} text="Close" />
					</DialogFooter>
				</Dialog>
				{snippetItems.length > 0 ?
					<h2 className="ms-font-m ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{useText}</h2>
					: null}
				<ul className="ms-List ms-welcome__features ms-u-slideUpIn10">
					{snippetItems}
				</ul>
				{
					this.state.isLoading ?
						<Spinner size={SpinnerSize.large} className="loading-spinner" />
						: null
				}
			</main>
		);
	}

}