import * as React from 'react';
import { Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';

export interface IDeleteConfirmDialogProps {
  itemTitle: string;
  onConfirm: () => void;
  onDismiss: () => void;
}

export const DeleteConfirmDialog: React.FC<IDeleteConfirmDialogProps> = ({ itemTitle, onConfirm, onDismiss }) => {
  return (
    <Dialog
      hidden={false}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: 'Delete Item',
        subText: `Are you sure you want to delete "${itemTitle}"? This action cannot be undone.`
      }}
      modalProps={{ isBlocking: true }}
    >
      <DialogFooter>
        <PrimaryButton text="Delete" onClick={onConfirm} />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};

export default DeleteConfirmDialog;
