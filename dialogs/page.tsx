// OutlookAddinIntegrated.tsx
'use client';
import React, { useEffect, useState } from 'react';
import { Checkbox, PrimaryButton, Stack, Text } from '@fluentui/react';
import PostsTab from '../taskpane/page';

export default function OutlookAddinIntegrated() {
  const [attachmentsFound, setAttachmentsFound] = useState(0);
  const [item, setItem] = useState(null);
  const [attachments, setAttachments] = useState([]);
  const [selectedAttachments, setSelectedAttachments] = useState<string[]>([]);
  const [showPostTab, setShowPostTab] = useState(false);

  const getAttachmentContent = () => {
    const item = Office.context.mailbox.item;

    const mailAttachments = item.attachments;
    if (mailAttachments.length <= 0) {
      console.log('Mail item has no attachments.');
      return;
    }

    setAttachments(mailAttachments);
    setAttachmentsFound(mailAttachments.length);
  };

  const handleAttachmentsCallback = (result) => {
    switch (result.value.format) {
      case Office.MailboxEnums.AttachmentContentFormat.Base64:
        console.log('Attachment is a Base64-encoded string.');
        break;
      case Office.MailboxEnums.AttachmentContentFormat.Eml:
        console.log('Attachment is a message.');
        break;
      case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
        console.log('Attachment is a calendar item.');
        break;
      case Office.MailboxEnums.AttachmentContentFormat.Url:
        console.log('Attachment is a cloud attachment.');
        break;
      default:
      // Handle attachment formats that aren't supported.
    }
  };

  const handleCheckboxChange = (attachmentId: string) => {
    setSelectedAttachments((prevSelection) =>
      prevSelection.includes(attachmentId)
        ? prevSelection.filter((id) => id !== attachmentId)
        : [...prevSelection, attachmentId]
    );
  };

  const handleSaveClick = () => {
    console.log('Selected Attachments:', selectedAttachments);
    setShowPostTab(true);
  };

  useEffect(() => {
    if (window.Office) {
      window.Office.onReady((info) => {
        if (info.host === Office.HostType.Outlook) {
          if (window.Office.context && window.Office.context.mailbox) {
            const mailboxItem = window.Office.context.mailbox.item;

            if (mailboxItem) {
              setItem(mailboxItem);
            } else {
              console.error('Office.context.mailbox.item is undefined.');
            }
          } else {
            console.error('Office.context.mailbox is undefined.');
          }
        }
      });
    }
  }, []);

  return (
    <div>
      <script type='text/javascript' src='https://appsforoffice.microsoft.com/lib/1/hosted/office.js'></script>
      <div style={{ borderBottom: '1px solid #ccc', paddingBottom: '10px', marginBottom: '10px' }}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <img
            src='https://images.g2crowd.com/uploads/product/image/small_square/small_square_74a710db08d54938cdebca50343049eb/cloudfiles-cloudfiles.png'
            alt='CloudFiles Logo'
            style={{ width: '24px', height: '24px', marginRight: '8px' }}
          />
          <h1 style={{ fontSize: '20px', fontWeight: 'bold', margin: '0' }}>CloudFiles</h1>
        </div>
      </div>

      <div style={{ borderBottom: '1px solid #ccc', paddingBottom: '10px', marginBottom: '10px' }}>
        <h2>Attachments Found: {attachmentsFound}</h2>
      </div>

      {/* Display Attachment Selector */}
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 20 } }}>
        <Text variant='xxLarge'>Select Attachments</Text>
        <Stack tokens={{ childrenGap: 10 }}>
          {attachments.map((attachment) => (
            <Checkbox
              key={attachment.id}
              label={attachment.name}
              checked={selectedAttachments.includes(attachment.id)}
              onChange={() => handleCheckboxChange(attachment.id)}
            />
          ))}
        </Stack>
        <PrimaryButton text='Save Selected Attachments' onClick={handleSaveClick} />
      </Stack>

      <div>
        <button style={{ padding: '10px', fontSize: '16px', marginLeft: '10px' }} onClick={getAttachmentContent}>
          Get Attachments
        </button>

        {showPostTab && <PostsTab contentType='folder' formType='Read' selectedAttachments={selectedAttachments} />}
      </div>
    </div>
  );
}
