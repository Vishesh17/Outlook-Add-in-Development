'use client';

import * as React from 'react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { useState, useEffect } from 'react';
import { DetailsRow } from '@fluentui/react/lib/DetailsList';
import { FluentProvider, Button, webLightTheme } from '@fluentui/react-components';
import { Icon } from '@fluentui/react/lib/Icon';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import FileDetails from './fileDetails';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

initializeIcons();

export const MyIcon = () => <Icon iconName='CircleAddition' />;
const BackIcon = () => <Icon iconName='Back' />;

export interface Payload {
  name: string;
  resource: string;
  resourceType: string;
  settings: { general: { allowDownloading: boolean } };
}

export interface IDocument {
  id: string;
  name: string;
  views: string;
  iconName: string;
  type: string;
  createdAt: string;
  updatedAt: string;
  updatedAtvalue: number;
  fileSize: string;
  fileSizeRaw: number;
}

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px',
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden',
      },
    },
    border: '1px solid #ddd',
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px',
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
  },
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px',
  },
};

const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { width: 200 } };

const uploadFile = async (url: string, name: string) => {
  Office.context.mailbox.item.body.getAsync('html', (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;
      const linkHtml = `<a href="${url}">${name}</a>`;
      const newEmailBody = emailBody + linkHtml;

      Office.context.mailbox.item.body.setAsync(newEmailBody, { coercionType: 'html' }, (setResult) => {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Email body updated successfully.');
        } else {
          console.error('Failed to set email body:', setResult.error);
        }
      });
    } else {
      console.error('Failed to get email body:', result.error);
    }
  });
};

export default function PostsTab({ contentType, formType, selectedAttachments }) {
  const headers = new Headers();
  headers.append('Authorization', 'Bearer X4Hrs7C7oIH4YOjoMvaeihSmA3ZkM5U9znjCRDtD');
  headers.append('x-external-user', 'gupta.vishesh@cloudfiles.io');

  const apiUrl = 'https://api.cloudfiles.dev/api';
  const initialFolderId = 'root';
  const [currentFolderId, setCurrentFolderId] = useState(initialFolderId);
  const [apiData, setApiData] = useState([]);
  const [folderHistory, setFolderHistory] = useState([initialFolderId]);
  const [searchValue, setSearchValue] = useState('');
  const [selectedFile, setSelectedFile] = useState(null);
  const [showFileDetails, setShowFileDetails] = useState(false);
  const [allowDownloading, setAllowDownloading] = React.useState(true);
  const [selectedDriveId, setSelectedDriveId] = useState('workspace');
  const [selectedLibrary, setSelectedLibrary] = useState('cloudfiles');
  const [libraryOptions, setLibraryOptions] = useState([]);
  const [driveIdOptions, setDriveIdOptions] = useState([]);

  const [attachments, setAttachments] = useState([]);

  const saveAttachments = async () => {
    try {
      const item = Office.context.mailbox.item;

      // Check if the item has attachments
      if (item.attachments && item.attachments.length > 0) {
        const attachments = item.attachments;

        for (let i = 0; i < attachments.length; i++) {
          // Log the attachment type and its contents to the console.
          item.getAttachmentContentAsync(attachments[i].id, handleAttachmentsCallback);
        }
        setAttachments(attachments);
      } else {
        console.log('Mail item has no attachments.');
      }
    } catch (error) {
      console.error('Error:', error);
    }
  };

  function handleAttachmentsCallback(result) {
    // Identifies whether the attachment is a Base64-encoded string, .eml file, .icalendar file, or a URL.
    switch (result.value.format) {
      case Office.MailboxEnums.AttachmentContentFormat.Base64:
        // Handle file attachment.
        console.log('Attachment is a Base64-encoded string.');
        break;
      case Office.MailboxEnums.AttachmentContentFormat.Eml:
        // Handle email item attachment.
        console.log('Attachment is a message.');
        break;
      case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
        // Handle .icalender attachment.
        console.log('Attachment is a calendar item.');
        break;
      case Office.MailboxEnums.AttachmentContentFormat.Url:
        // Handle cloud attachment.
        console.log('Attachment is a cloud attachment.');
        break;
      default:
      // Handle attachment formats that aren't supported.
    }
  }

  const getlibraries = () => {
    const libraryurl = 'https://api.cloudfiles.dev/api/users/libraries';
    return fetch(libraryurl, { method: 'GET', headers: headers })
      .then((response) => response.json())
      .then((result) => {
        setLibraryOptions(result.map((library) => ({ key: library.app, text: library.name })));

        // Only set the selected library if it's not already set

        if (!selectedLibrary) {
          const defaultLibrary = result.length > 0 ? result[0].app : null;
          setSelectedLibrary(defaultLibrary);
        }

        const selectedLibraryData = result.find((library) => library.app === selectedLibrary);
        if (selectedLibraryData) {
          setDriveIdOptions(
            selectedLibraryData.labels.map((label) => ({ key: label.id, text: label.name, type: label.type }))
          );
          const defaultDrive = selectedLibraryData.labels.length > 0 ? selectedLibraryData.labels[0].id : null;
          setSelectedDriveId(defaultDrive);
        }
        return result;
      })
      .catch((error) => {
        console.log('Error: ', error);
        return null;
      });
  };

  // Usage
  useEffect(() => {
    const fetchLibraries = async () => {
      const response = await getlibraries();
      if (response) {
        console.log('Full response:', response);

        // Set the default drive only if it's not already set
        if (!selectedDriveId) {
          const selectedLibraryData = response.find((library) => library.app === selectedLibrary);
          if (selectedLibraryData) {
            const defaultDrive = selectedLibraryData.labels.length > 0 ? selectedLibraryData.labels[0].id : null;
            setSelectedDriveId(defaultDrive);
          }
        }
      }
    };

    fetchLibraries();
  }, [selectedLibrary]);

  console.log(selectedDriveId, selectedLibrary);

  useEffect(() => {
    let fetchUrl;

    if (selectedLibrary === 'sharepoint') {
      fetchUrl = `${apiUrl}/sites/${selectedDriveId}/children?contentType=${contentType}&library=${selectedLibrary}&limit=100`;
    } else if (selectedDriveId) {
      fetchUrl = `${apiUrl}/drives/${selectedDriveId}/children?contentType=${contentType}&library=${selectedLibrary}&limit=100`;
    }

    fetch(fetchUrl, { method: 'GET', headers: headers })
      .then((response) => response.json())
      .then((result) => {
        if (Array.isArray(result.content)) {
          const filteredData = result.content.filter((item) =>
            item.name.toLowerCase().includes(searchValue.toLowerCase())
          );
          setApiData(filteredData);
          console.log(apiData);
        } else {
          console.error('API response is not an array:', result);
        }
      })
      .catch((error) => {
        console.error('Error:', error);
      });
  }, [currentFolderId, searchValue, selectedLibrary, selectedDriveId]);

  const handleBackClick = () => {
    // If the current folder is not the root, go back to the parent folder
    if (selectedFile && folderHistory.length > 1) {
      const previousFolderId = folderHistory[folderHistory.length - 1];
      setFolderHistory(folderHistory.slice(0, -1)); // Remove the last folder from history
      setCurrentFolderId(previousFolderId);
      setSelectedFile(null);
    } else if (folderHistory.length > 1) {
      // If a file is not selected, go back to the parent folder
      folderHistory.pop(); // Remove the last folder from history
      const previousFolderId = folderHistory[folderHistory.length - 1];
      setCurrentFolderId(previousFolderId);
      setFolderHistory([...folderHistory]);
      setSelectedFile(null);
    }
  };

  const handleClick = async (folderId: string, fileType: string) => {
    if (fileType === 'folder') {
      const folderUrl = `${apiUrl}/folders/${folderId}/children?contentType=${contentType}&driveId=${selectedDriveId}&library=${selectedLibrary}&limit=100`;

      fetch(folderUrl, { method: 'GET', headers: headers })
        .then((response) => response.json())
        .then((result) => {
          if (Array.isArray(result.content)) {
            const filteredData = result.content.filter((item) =>
              item.name.toLowerCase().includes(searchValue.toLowerCase())
            );
            setApiData(filteredData);
          } else {
            console.error('API response is not an array:', result);
          }
        })
        .catch((error) => {
          console.error('Error:', error);
        });

      setFolderHistory([...folderHistory, folderId]);
      setCurrentFolderId(folderId);
      setSelectedFile(null);
    } else {
      // Handle file click
      const clickedFile = apiData.find((file) => file.id === folderId);
      setFolderHistory([...folderHistory, currentFolderId]);
      setCurrentFolderId(clickedFile.parentId);
      setSelectedFile(clickedFile);
      console.log(clickedFile);
      setShowFileDetails(true);
    }
  };

  const getLinkUrl = async (payload: Payload, { driveId, library }) => {
    const fileUrl = `https://api.cloudfiles.dev/api/links?library=${library}&driveId=${driveId}`;

    const headers = new Headers();
    headers.append('Content-Type', 'application/json');
    headers.append('Authorization', 'Bearer X4Hrs7C7oIH4YOjoMvaeihSmA3ZkM5U9znjCRDtD');
    headers.append('x-external-user', 'gupta.vishesh@cloudfiles.io');

    fetch(fileUrl, { method: 'POST', headers: headers, body: JSON.stringify(payload) })
      .then((response) => response.json())
      .then((result) => {
        const url = result.url;
        const name = payload.name;
        console.log('Name: ', name);
        console.log('URL : ', url);
        uploadFile(url, name);
        return url;
      })
      .catch((error) => {
        console.error('Error:', error);
      });
  };

  const handleInsertFileClick = (payload: Payload, { driveId, library }) => {
    getLinkUrl(payload, { driveId, library }); // Call the getLinkUrl function
  };

  const columns: IColumn[] = [
    {
      key: 'column1',
      name: 'Name',
      fieldName: 'name',
      minWidth: 70,
      maxWidth: 220,
      isResizable: true,
      onRender: (item, index) => (
        <div style={{ display: 'flex', alignItems: 'center' }} onClick={() => handleClick(item.id, item.type)}>
          {item.type === 'folder' ? (
            <Icon iconName='FabricFolder' style={{ marginRight: '8px' }} /> // Use folder icon if it's a folder
          ) : (
            <Icon iconName='Page' style={{ marginRight: '8px' }} /> // Use document icon if it's a file
          )}
          {item.name}
        </div>
      ),
    },
    {
      key: 'column2',
      name: '',
      fieldName: 'toolbar',
      minWidth: 20,
      maxWidth: 30,
      isResizable: true,
      onRender: (item, index, column) => (
        <FluentProvider theme={webLightTheme}>
          <Button
            onClick={() =>
              getLinkUrl(
                {
                  name: item.name,
                  resource: item.id,
                  resourceType: item.type,
                  settings: { general: { allowDownloading: allowDownloading } },
                },
                { driveId: selectedDriveId, library: selectedLibrary }
              )
            }
            appearance='transparent'
            icon={<MyIcon />}
          ></Button>
        </FluentProvider>
      ),
    }, // Add more columns as needed
  ];
  const onRenderRow = (props) => {
    const customStyles = {};
    return <DetailsRow {...props} styles={customStyles} />;
  };

  const _selection = new Selection({
    onSelectionChanged: () => {
      // Handle selection changes if needed
    },
  });

  const items = apiData.map((item) => ({
    key: item.id,
    name: item.name,
    type: item.type,
    // Add more properties as needed
  }));

  const onActiveItemChanged = (item: any, index: number, ev: React.FocusEvent<HTMLElement>): void => {
    // Handle active item changed
    _selection.setAllSelected(false);
    _selection.toggleIndexSelected(index);
  };

  function openDialog(url) {
    Office.context.ui.displayDialogAsync(url, { height: 50, width: 50 }, function (result) {
      var dialog = result.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        // Check if the received event has the 'message' property
        if ('message' in arg) {
          var messageFromDialog = arg.message;
          console.log('Message from dialog: ' + messageFromDialog);
        } else {
          // Handle the case where 'message' property is not present (optional)
          console.warn("Unexpected event format. No 'message' property found.");
        }
      });
    });
  }

  return (
    <div style={{ padding: 15 }}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: 10, justifyContent: 'space-between' }}>
        {formType !== 'Read' && (
          <div style={{ display: 'flex', alignItems: 'center', marginBottom: 10, justifyContent: 'space-between' }}>
            <img
              src='https://images.g2crowd.com/uploads/product/image/small_square/small_square_74a710db08d54938cdebca50343049eb/cloudfiles-cloudfiles.png'
              alt='CloudFiles Logo'
              style={{ width: '24px', height: '24px', marginRight: '8px' }}
            />
            <h1 style={{ fontSize: '20px', fontWeight: 'bold' }}>CloudFiles</h1>
          </div>
        )}

        <div style={{ display: 'flex', alignItems: 'center', marginBottom: 10, justifyContent: 'space-between' }}>
          {currentFolderId !== 'root' && (
            <Button
              shape='square'
              style={{ padding: '1px', border: 'none', cursor: 'pointer' }}
              onClick={handleBackClick}
              disabled={currentFolderId === 'root' && !selectedFile}
              icon={<BackIcon />}
            ></Button>
          )}
        </div>
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', marginBottom: 10 }}>
        <div>
          <Dropdown
            label='Select Library'
            selectedKey={selectedLibrary}
            options={libraryOptions}
            onChange={(event, option) => {
              setSelectedLibrary(option.key as string);

              // Assuming you have the selectedLibraryData available
              const selectedLibraryData = libraryOptions.find((library) => library.app === option.key);

              if (selectedLibraryData) {
                setDriveIdOptions(
                  selectedLibraryData.labels.map((label) => ({ key: label.id, text: label.name, type: label.type }))
                );
              }
            }}
            style={{ width: '370px' }} // Adjust the width as needed
          />
        </div>
        <div style={{ marginBottom: 10 }}>
          <Dropdown
            label='Select Drive'
            selectedKey={selectedDriveId}
            options={driveIdOptions}
            onChange={(event, option) => setSelectedDriveId(option.key as string)}
            style={{ width: '370px' }} // Adjust the width as needed
          />
        </div>
      </div>
      <div style={{ padding: '10px 10px 20px 10px' }}>
        <SearchBox
          placeholder='Search'
          onChange={(event, newValue) => setSearchValue(newValue)}
          value={searchValue}
          iconProps={{ iconName: 'Search' }}
        />

        {showFileDetails && selectedFile ? (
          <FileDetails
            fileName={selectedFile?.name || 'No Name'} // Use a default name if 'name' is undefined
            allowDownloading={allowDownloading}
            onInsertFileClick={(payload: Payload) => {
              // Implement your logic here
              handleInsertFileClick(
                {
                  name: selectedFile?.name,
                  resource: selectedFile?.id,
                  resourceType: selectedFile?.type,
                  settings: { general: { allowDownloading: allowDownloading } },
                },
                { driveId: selectedDriveId, library: selectedLibrary }
              );
            }}
            onPreviewClick={(payload) => {
              const url = `https://app.cloudfiles.dev/preview/file/${selectedFile.id}?library=${selectedLibrary}&driveId=${selectedDriveId}?back=true`;
              openDialog(url);
              // Implement your logic here
              console.log('Preview clicked');
            }}
            onAllowDownloadChange={(allowDownloading) => setAllowDownloading(allowDownloading)}
          />
        ) : (
          <div>
            {apiData.length > 0 ? (
              <DetailsList
                items={apiData}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                onRenderRow={onRenderRow}
                onActiveItemChanged={onActiveItemChanged}
                styles={{
                  root: {
                    borderBottom: '1px solid #ccc', // Add a border between rows
                  },
                  headerWrapper: {
                    borderBottom: '1px solid #ccc', // Add a border between header and rows
                  },
                  header: {
                    selectors: {
                      ':not(:last-child)': {
                        borderRight: '1px solid #ccc', // Add a border between columns
                      },
                    },
                  },
                  focusZone: {
                    selectors: {
                      ':not(:last-child)': {
                        borderRight: '1px solid #ccc', // Add a border between columns
                      },
                    },
                  },
                  cell: {
                    selectors: {
                      ':not(:last-child)': {
                        borderRight: '2px solid #ccc', // Add a border between columns
                      },
                    },
                  },
                }}
              />
            ) : (
              <div style={{ textAlign: 'center' }}>
                <img
                  src={
                    'https://www.google.com/imgres?imgurl=https%3A%2F%2Fcdn-icons-png.flaticon.com%2F512%2F2739%2F2739782.png&tbnid=kXSoZXFpfU40NM&vet=12ahUKEwim_sDzoIyDAxXufGwGHVRvAv4QMygOegQIARB_..i&imgrefurl=https%3A%2F%2Fwww.flaticon.com%2Ffree-icon%2Fempty-folder_2739782&docid=trFVJh1zzkAxpM&w=512&h=512&q=empty%20folder%20icon&ved=2ahUKEwim_sDzoIyDAxXufGwGHVRvAv4QMygOegQIARB_'
                  }
                  alt='Empty Folder'
                  style={{ maxWidth: '100%', maxHeight: '100%' }}
                />
                <p>This folder is empty.</p>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
