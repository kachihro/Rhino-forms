import * as React from 'react';
import { Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';

export interface IDeleteConfirmDialogProps {
  itemTitle: string;
  itemLabel?: string;
  busy?: boolean;
  onConfirm: () => void;
  onDismiss: () => void;
}

export const DeleteConfirmDialog: React.FC<IDeleteConfirmDialogProps> = ({ itemTitle, itemLabel, busy, onConfirm, onDismiss }) => (
  <Dialog
    hidden={false}
    onDismiss={onDismiss}
    dialogContentProps={{
      type: DialogType.normal,
      title: `Delete ${itemLabel || 'item'}?`,
      subText: `"${itemTitle}" will be moved to the site recycle bin.`
    }}
    modalProps={{ isBlocking: true }}
  >
    <DialogFooter>
      <PrimaryButton text={busy ? 'Deleting…' : 'Delete'} onClick={onConfirm} disabled={busy}
        styles={{ root: { background: '#c50f1f', borderColor: '#c50f1f' }, rootHovered: { background: '#a4262c', borderColor: '#a4262c' } }} />
      <DefaultButton text="Cancel" onClick={onDismiss} disabled={busy} />
    </DialogFooter>
  </Dialog>
);

export default DeleteConfirmDialog;
