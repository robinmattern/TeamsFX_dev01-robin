import React from 'react';
import './App.css';
import { TeamsProvider } from '@microsoft/mgt-teams-provider';
import { Providers } from '@microsoft/mgt-element';
import * as microsoftTeams from '@microsoft/teams-js';
import { IDynamicPerson, Login, PeoplePicker, PersonType } from '@microsoft/mgt-react';
import {
  Button,
  Flex,
  FlexItem,
  Text,
  Provider,
  Segment,
  TextArea,
  Dropdown,
  Dialog,
  Loader,
  CloseIcon,
} from '@fluentui/react-northstar';
import { useTeams } from 'msteams-react-base-component';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { useConfiguration } from './hooks/useConfiguration';
import { useGraphService } from './hooks/useGraphService';
import { MgtPerson } from '@microsoft/mgt';

export const App: React.FunctionComponent = () => {
  const [{ inTeams, theme, themeString }] = useTeams({});
  const [isSignedIn] = useIsSignedIn();
  const [config] = useConfiguration();
  const [{ loading, sendNotification }] = useGraphService();
  const [userId, setUserId] = React.useState<string>('');
  const [people, setPeople] = React.useState<IDynamicPerson[]>([]);
  const [message, setMessage] = React.useState<string>('');
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const [mgtTheme, setMgtTheme] = React.useState<string>('');
  const [dialogOpen, setDialogOpen] = React.useState<boolean>(false);
  const [category, setCategory] = React.useState<string | null>(null);

  React.useEffect(() => {
    TeamsProvider.microsoftTeamsLib = microsoftTeams;
    Providers.globalProvider = new TeamsProvider({
      clientId: config.clientId,
      authPopupUrl: '/auth.html',
      scopes: ['User.Read', 'User.ReadBasic.All', 'People.Read', 'TeamsActivity.Send', 'AppCatalog.Read.All'],
    });
    MgtPerson.config.useContactApis = false;
  }, [config.clientId]);

  React.useEffect(() => {
    if (themeString === 'default') {
      setMgtTheme('mgt-light');
    } else if (themeString === 'dark') {
      setMgtTheme('mgt-dark');
    }
  }, [themeString]);

  const _sendNotification = async () => {
    if (isSignedIn) {
      try {
        setDialogOpen(true);
        await sendNotification(config.teamsAppId, userId, category, 'KudosHome', 'sendKudosToUser', message);
        setDialogOpen(false);
        _clearForm();
      } catch (e) {
        setErrorMessage(e.message);
      }
    }
  };

  const _clearForm = () => {
    setUserId('');
    setPeople([]);
    setCategory(null);
    setMessage('');
  };

  const _onPeopleChanged = (event: any) => {
    if (event.detail.length) {
      setUserId(event.detail[0].id);
      setPeople(event.detail);
    } else {
      setUserId('');
      setPeople([]);
    }
  };

  const _onMessageChanged = (event: any, component: any) => {
    setMessage(component.value);
  };

  const _onCategoryChanged = (event: any, component: any) => {
    setCategory(component.value);
  };

  const _isDisabled = (): boolean | undefined => {
    return message === '' || userId === '' || category === '';
  };

  const _closeDialog = () => {
    setDialogOpen(false);
    setErrorMessage('');
  };

  const formStyles: any = {
    paddingBottom: '30px',
  };
  return (
    <>
      {inTeams && (
        <Provider theme={theme}>
          <div className={mgtTheme}>
            <Flex>
              <FlexItem push styles={{ padding: '0', border: '0', boxShadow: '0' }}>
                <Segment content="Logout">
                  <Login></Login>
                </Segment>
              </FlexItem>
            </Flex>

            {isSignedIn && (
              <Flex gap="gap.small" padding="padding.medium">
                <Flex.Item grow={true}>
                  <Flex column>
                    <Text content="Use this form to send Kudos to your colleagues!" styles={formStyles} />
                    <PeoplePicker
                      type={PersonType.person}
                      selectionMode="single"
                      placeholder="Search for a colleague"
                      selectionChanged={_onPeopleChanged}
                      selectedPeople={people}
                    ></PeoplePicker>
                    <Dropdown
                      fluid
                      items={[
                        'Thank you',
                        'Customer advocate',
                        'Growth mindset',
                        'Team player',
                        'New perspective',
                        'Extra mile',
                        'Leadership',
                      ]}
                      placeholder="Select the type of kudos"
                      styles={{ marginTop: '15px' }}
                      onChange={_onCategoryChanged}
                      value={category}
                    />
                    <TextArea
                      fluid
                      placeholder="Short message"
                      onChange={_onMessageChanged}
                      maxLength={128}
                      styles={{ height: '100px', marginBottom: '15px', marginTop: '15px' }}
                      value={message}
                    />
                    <Button content="Submit" onClick={_sendNotification} disabled={_isDisabled() || loading} />

                    <Dialog
                      header="Sending your Kudos"
                      open={dialogOpen}
                      headerAction={{
                        icon: <CloseIcon />,
                        title: 'Close',
                        onClick: _closeDialog,
                      }}
                      content={
                        <>
                          {!errorMessage && <Loader />}
                          {errorMessage && <Text important content={errorMessage} />}
                        </>
                      }
                    />
                  </Flex>
                </Flex.Item>
              </Flex>
            )}
          </div>
        </Provider>
      )}
    </>
  );
};
