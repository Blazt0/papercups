import React from 'react';
import {RouteComponentProps} from 'react-router';
import {Link} from 'react-router-dom';
import {Box, Flex} from 'theme-ui';
import qs from 'query-string';
import {
  notification,
  Alert,
  Button,
  Card,
  Container,
  Divider,
  Paragraph,
  Popconfirm,
  Tag,
  Text,
  Title,
} from '../common';
import {ArrowLeftOutlined, CheckCircleOutlined} from '../icons';
import * as API from '../../api';
import {Account, SlackAuthorization} from '../../types';
import {getSlackAuthUrl, getSlackRedirectUrl} from './support';
import logger from '../../logger';

type Props = RouteComponentProps<{}>;
type State = {
  status: 'loading' | 'success' | 'error';
  authorization: SlackAuthorization | null;
  account: Account | null;
  error: any;
};

class SlackReplyIntegrationDetails extends React.Component<Props, State> {
  state: State = {
    status: 'loading',
    account: null,
    authorization: null,
    error: null,
  };

  async componentDidMount() {
    try {
      const {location, history} = this.props;
      const {search} = location;
      const q = qs.parse(search);
      const code = q.code ? String(q.code) : null;

      if (code) {
        await this.authorize(code, q);

        history.push(`/integrations/slack/reply`);
      }

      this.fetchSlackAuthorization();
    } catch (error) {
      logger.error(error);

      this.setState({status: 'error', error});
    }
  }

  fetchSlackAuthorization = async () => {
    try {
      const auth = await API.fetchSlackAuthorization('reply');
      const account = await API.fetchAccountInfo();

      this.setState({account, authorization: auth, status: 'success'});
    } catch (error) {
      logger.error(error);

      this.setState({status: 'error', error});
    }
  };

  authorize = async (code: string, query: any) => {
    const state = query.state ? String(query.state) : null;

    if (!code) {
      return null;
    }

    const authorizationType = state || 'reply';

    return API.authorizeSlackIntegration({
      code,
      type: authorizationType,
      redirect_url: getSlackRedirectUrl(),
    })
      .then((result) => logger.debug('Successfully authorized Slack:', result))
      .catch((err) => {
        logger.error('Failed to authorize Slack:', err);

        const description =
          err?.response?.body?.error?.message || err?.message || String(err);

        notification.error({
          message: 'Failed to authorize Slack',
          duration: null,
          description,
        });
      });
  };

  disconnect = () => {
    const {authorization} = this.state;
    const authorizationId = authorization?.id;

    if (!authorizationId) {
      return null;
    }

    return API.deleteSlackAuthorization(authorizationId)
      .then(() => this.fetchSlackAuthorization())
      .catch((err) =>
        logger.error('Failed to remove Slack authorization:', err)
      );
  };

  isOnStarterPlan = () => {
    const {account} = this.state;

    if (!account) {
      return false;
    }

    return account.subscription_plan === 'starter';
  };

  render() {
    const {authorization, status} = this.state;

    if (status === 'loading') {
      return null;
    }

    const hasAuthorization = !!(authorization && authorization.id);

    return (
      <Container sx={{maxWidth: 720}}>
        <Box mb={4}>
          <Link to="/integrations">
            <Button icon={<ArrowLeftOutlined />}>Back to integrations</Button>
          </Link>
        </Box>

        {this.isOnStarterPlan() && (
          <Box mb={4}>
            <Alert
              message={
                <Text>
                  This integration is only available on the Lite and Team
                  subscription plans.{' '}
                  <Link to="billing">Sign up for a free trial!</Link>
                </Text>
              }
              type="warning"
              showIcon
            />
          </Box>
        )}

        <Box mb={4}>
          <Title level={3}>Reply from Slack</Title>

          <Paragraph>
            <Text>
              Reply to messages from your customers directly through Slack.
            </Text>
          </Paragraph>
        </Box>

        <Box mb={4}>
          <Card sx={{p: 3}}>
            <Paragraph>
              <Flex sx={{alignItems: 'center'}}>
                <img src="/slack.svg" alt="Slack" style={{height: 20}} />
                <Text strong style={{marginLeft: 8}}>
                  How does it work?
                </Text>
              </Flex>
            </Paragraph>

            <Text type="secondary">
              When you link Papercups with Slack, all new incoming messages will
              be forwarded to the Slack channel of your choosing. From the
              comfort of your team's Slack workspace, you can reply to and
              resolve conversations with your users.
            </Text>
          </Card>
        </Box>

        <Divider />

        <Box mb={4}>
          <Card sx={{p: 3}}>
            <Flex sx={{justifyContent: 'space-between'}}>
              <Flex sx={{alignItems: 'center'}}>
                <img src="/slack.svg" alt="Slack" style={{height: 20}} />
                <Text strong style={{marginLeft: 8, marginRight: 8}}>
                  Reply to messages from Slack
                </Text>
                {hasAuthorization && (
                  <Tag icon={<CheckCircleOutlined />} color="success">
                    connected
                  </Tag>
                )}
              </Flex>

              {hasAuthorization ? (
                <Flex mx={-1}>
                  <Box mx={1}>
                    <a href={getSlackAuthUrl('reply')}>
                      <Button>Reconnect</Button>
                    </a>
                  </Box>
                  <Box mx={1}>
                    <Popconfirm
                      title="Are you sure you want to disconnect from Slack?"
                      okText="Yes"
                      cancelText="No"
                      placement="topLeft"
                      onConfirm={() => this.disconnect()}
                    >
                      <Button type="primary" danger>
                        Disconnect
                      </Button>
                    </Popconfirm>
                  </Box>
                </Flex>
              ) : (
                <a href={getSlackAuthUrl('reply')}>
                  <Button type="primary">Connect</Button>
                </a>
              )}
            </Flex>
          </Card>
        </Box>

        <Box mb={4}>
          <Title level={4}>Integration settings</Title>

          <Card
            sx={{
              p: 3,
              bg: 'rgb(245, 245, 245)',
              opacity: hasAuthorization ? 1 : 0.6,
            }}
          >
            <Box mb={3}>
              <label>Channel</label>
              <Box>
                {authorization && authorization.channel ? (
                  <Text strong>{authorization.channel}</Text>
                ) : (
                  <Text type="secondary">Not connected</Text>
                )}
              </Box>
            </Box>
            <Box mb={3}>
              <label>Team</label>
              <Box>
                {authorization && authorization.team_name ? (
                  <Text strong>{authorization.team_name}</Text>
                ) : (
                  <Text type="secondary">Not connected</Text>
                )}
              </Box>
            </Box>
            <Box mb={3}>
              <label>Slack configuration URL</label>
              <Box>
                {authorization && authorization.configuration_url ? (
                  <Text>
                    <a
                      href={authorization.configuration_url}
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      {authorization.configuration_url}
                    </a>
                  </Text>
                ) : (
                  <Text type="secondary">Not connected</Text>
                )}
              </Box>
            </Box>
          </Card>
        </Box>
      </Container>
    );
  }
}

export default SlackReplyIntegrationDetails;
