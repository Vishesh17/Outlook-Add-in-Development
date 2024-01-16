'use client';
// FileDetails.tsx
import * as React from 'react';
import { Button, makeStyles } from '@fluentui/react-components';
import { Icon } from '@fluentui/react/lib/Icon';
import { Switch } from '@fluentui/react-components';
import { Payload } from './page';

interface FileDetailsProps {
  fileName: string;
  allowDownloading: boolean;
  onInsertFileClick: (item: any) => void;
  onPreviewClick: (item: any) => void;
  onAllowDownloadChange: (allowDownloading: boolean) => void;
}

const MyIcon = () => <Icon iconName='RedEye' />;

const useStyles = makeStyles({
  wrapper: {
    columnGap: '15px',
    display: 'flex',
  },
});

const FileDetails: React.FC<FileDetailsProps> = ({ fileName, onInsertFileClick, onPreviewClick }) => {
  const styles = useStyles();
  const [checked, setChecked] = React.useState(true);
  const onChange = React.useCallback(
    (ev) => {
      setChecked(ev.currentTarget.checked);
    },
    [setChecked]
  );

  return (
    <div style={{ padding: 20 }}>
      <h1 style={{ fontSize: '18px', fontWeight: 'bold' }}>
        {' '}
        <Icon iconName='Page' style={{ marginRight: '8px' }} />
        {fileName}
      </h1>
      <div style={{ padding: 10 }}>
        <div
          style={{
            padding: 5,
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: 10,
            display: 'flex',
          }}
        >
          <Switch
            checked={checked}
            onChange={onChange}
            label={checked ? 'Allow Downloading' : 'Disallow Downloading'}
          />
          <Button icon={<MyIcon />} onClick={onPreviewClick}></Button>
        </div>
        <div style={{ padding: 5 }}>
          <Button style={{ fontSize: '12px', fontWeight: 'bold' }} appearance='primary' onClick={onInsertFileClick}>
            Insert File
          </Button>
        </div>
      </div>
    </div>
  );
};

export default FileDetails;
