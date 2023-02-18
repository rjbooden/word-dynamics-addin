/* eslint-disable prettier/prettier */
import { Dialog, DialogFooter, DialogType, IconButton, PrimaryButton, Spinner, SpinnerSize, Stack, TextField } from "@fluentui/react";
import React, { useState, useEffect } from "react";
import { strings } from "../../services/LocaleService";
import { SettingsService, Snippet } from "../../services/SettingsService";
import Wait from "../../helper/Wait";
import ClickableListItem from "./ClickableListItem";

export interface SnippetProps {
	onError: (errorMessage: string, errorInfo?: string) => void;
}

export default function SnippetsView(props: SnippetProps) {
	let isWorking: boolean = false;

	const [newSnippetName, setSnippetName] = useState<string>('');
	const [snippets, setSnippets] = useState<Snippet[]>([]);
	const [isLoading, setLoading] = useState<boolean>(true);
	const [isSaving, setSaving] = useState<boolean>(false);
	const [hideDialog, setHideDialog] = useState<boolean>(true);

	useEffect(() => {
		SettingsService.loadSnippets().then((snippets) => {
			setSnippets(snippets);
			setLoading(false);
		}).catch((reason) => {
			// eslint-disable-next-line no-undef
			console.log(reason);
			props.onError(strings.unableLoadingSnippets, reason);
			setLoading(false);
		});
	}, []);

	// eslint-disable-next-line no-undef
	const ensureFullyIncludedContentControls = async (context: Word.RequestContext, range: Word.Range) => {
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
	};

	const onCreateSnippet = () => {
		setSaving(true);
		isWorking = true;
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
					range = await ensureFullyIncludedContentControls(context, range);
					const ooxml = range.getOoxml();
					await context.sync();
					// eslint-disable-next-line office-addins/load-object-before-read
					let newSnippets: Snippet[] = JSON.parse(JSON.stringify(snippets));
					// eslint-disable-next-line office-addins/load-object-before-read
					snippets.push({ title: newSnippetName, content: ooxml.value });
					SettingsService.saveSnippets(newSnippets).then(() => {
						setSnippets(snippets);
						setSaving(false);
						setSnippetName('');
						props.onError(null);
						isWorking = false;
					}).catch((reason) => {
						// eslint-disable-next-line no-undef
						console.log(reason);
						props.onError(strings.unableStoringSnippet, reason);
						setSaving(false);
						isWorking = false;
					});
				}
				else {
					setSaving(false);
					setHideDialog(false);
					isWorking = false;
				}
			} catch (e) {
				// eslint-disable-next-line no-undef
				console.log(e);
				props.onError(e);
				setSaving(false);
				isWorking = false;
			}
		});
	}

	const closeDialog = () => {
		setHideDialog(true);
	}

	const deleteSnippet = (s: Snippet) => {
		const handleDelete = () => {
			isWorking = true;
			let deepCopy: Snippet[] = JSON.parse(JSON.stringify(snippets));
			let newSnippets = deepCopy.filter((item) => { return item.title != s.title });
			SettingsService.saveSnippets(newSnippets).then(() => {
				setSnippets(newSnippets);
				props.onError(null);
				isWorking = false;
			}).catch((reason) => {
				setSnippets(newSnippets); // to clear deleting spinner
				// eslint-disable-next-line no-undef
				console.log(reason);
				props.onError(strings.unableDeletingSnippet, reason);
				isWorking = false;
			});
		};
		if (isWorking) { // handle multiple deletes sequentially
			Wait(200, 80, () => {
				return isWorking;
			}).then(() => {
				handleDelete();
			}).catch((reason) => {
				// eslint-disable-next-line no-undef
				console.log(reason);
				props.onError(strings.unableDeletingSnippet, reason);
			});
		}
		else {
			handleDelete();
		}
	}

	const insertSnippet = (Snippet: Snippet, notifyLoaded) => {
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
				props.onError(e);
				notifyLoaded();
			}
		});
	}

	let snippetItems = snippets.map((s, i) => {
		return (<ClickableListItem
			label={s.title}
			iconName="Paste"
			key={s.title + i}
			showLoading={true}
			onDelete={() => { deleteSnippet(s); }}
			onClick={(notifyLoaded) => { insertSnippet(s, notifyLoaded); }} />);
	});

	const enableCopy = newSnippetName != ''
		&& (snippets.length === 0
			|| null == (snippets.find((s) => {
				return s.title == newSnippetName
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
				<TextField label={strings.newSnippetName} value={newSnippetName} className="snippet-name" onChange={(_e, v) => { setSnippetName(v) }} />
				{
					isSaving ?
						<Spinner size={SpinnerSize.small} className="saving-spinner" />
						: <IconButton
							iconProps={{ iconName: 'SaveTemplate' }}
							onClick={onCreateSnippet}
							disabled={!enableCopy}
							className="snippet-button"
						/>
				}
			</Stack>
			<Dialog
				hidden={hideDialog}
				onDismiss={closeDialog}
				dialogContentProps={dialogContentProps}
				modalProps={{ isBlocking: true }}
			>
				<DialogFooter>
					<PrimaryButton onClick={closeDialog} text="Close" />
				</DialogFooter>
			</Dialog>
			{snippetItems.length > 0 ?
				<h2 className="ms-font-m ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{useText}</h2>
				: null}
			<ul className="ms-List ms-welcome__features ms-u-slideUpIn10">
				{snippetItems}
			</ul>
			{
				isLoading ?
					<Spinner size={SpinnerSize.large} className="loading-spinner" />
					: null
			}
		</main>
	);
}

