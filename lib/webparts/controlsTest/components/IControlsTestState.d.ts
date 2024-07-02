/// <reference types="react" />
import { ImageSize } from "../../../FileTypeIcon";
import { IProgressAction } from "../../../Progress";
import { IFilePickerResult } from "../../../FilePicker";
import { ITag } from "@fluentui/react";
import { ITermInfo, ITermSetInfo, ITermStoreInfo } from "@pnp/sp/taxonomy";
export interface IControlsTestState {
    imgSize: ImageSize;
    items: any[];
    initialValues: any[];
    iFrameDialogOpened?: boolean;
    iFramePanelOpened?: boolean;
    authorEmails: string[];
    selectedList: string;
    progressActions: IProgressAction[];
    currentProgressActionIndex?: number;
    dateTimeValue: Date;
    richTextValue: string;
    currentCarouselElement: JSX.Element;
    canMovePrev: boolean;
    canMoveNext: boolean;
    comboBoxListItemPickerListId: string;
    comboBoxListItemPickerIds: any[];
    filePickerResult?: IFilePickerResult[];
    treeViewSelectedKeys?: string[];
    showAnimatedDialog?: boolean;
    showCustomisedAnimatedDialog?: boolean;
    showSuccessDialog?: boolean;
    showErrorDialog?: boolean;
    selectedTeam: ITag[];
    selectedTeamChannels: ITag[];
    filePickerDefaultFolderAbsolutePath?: string;
    errorMessage?: string;
    termPanelIsOpen?: boolean;
    actionTermId?: string;
    clickedActionTerm?: ITermInfo;
    selectedFilters?: string[];
    termStoreInfo: ITermStoreInfo;
    termSetInfo: ITermSetInfo;
    testTerms: ITermInfo[];
}
//# sourceMappingURL=IControlsTestState.d.ts.map