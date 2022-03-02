module.exports = function (RED) {
    'use strict';

    const {WebClient, LogLevel} = require('@slack/web-api');

    // TODO `retryConfig` (https://slack.dev/node-slack-sdk/web-api#automatic-retries)
    // TODO rate limiting (https://slack.dev/node-slack-sdk/web-api#rate-limits)
    // TODO `maxRequestConcurrency` (https://slack.dev/node-slack-sdk/web-api#request-concurrency)
    function WebClientConfig(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        const token = this.credentials.token;
        this.webClient = new WebClient(token, {
            logLevel: LogLevel.DEBUG,
            logger: {
                setLevel(level) {
                },
                setName(name) {
                },
                debug(...msg) {
                    node.debug(msg);
                },
                info(...msg) {
                    node.log(msg);
                },
                warn(...msg) {
                    node.warn(msg);
                },
                error(...msg) {
                    node.error(msg);
                }
            }
        });
    }

    function WebClientNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        const webClientConfig = RED.nodes.getNode(config.webClientConfig);
        this.webClient = webClientConfig.webClient;

        this.methodName = config.methodName;
        this.paginate = config.paginate;
        if (config.pageLimit) {
            this.pageLimit = parseInt(config.pageLimit);
        }
        if (config.shouldStopExpression) {
            this.shouldStopExpression = RED.util.prepareJSONataExpression(config.shouldStopExpression, this);
        }

        this.on('input', async function (msg, _send, done) {
            const methodName = node.methodName;
            const options = msg.payload || {};

            try {
                let ok = true;
                let msgs;

                if (node.paginate) {
                    // https://api.slack.com/docs/pagination
                    // https://slack.dev/node-slack-sdk/web-api#pagination
                    // TODO add support for classic pagination (https://api.slack.com/docs/pagination#classic)
                    options.limit = options.limit || node.pageLimit;

                    msgs = await node.webClient.paginate(methodName, options, function (page) {
                        if (node.shouldStopExpression) {
                            return RED.util.evaluateJSONataExpression(node.shouldStopExpression, page);
                        } else {
                            return false;
                        }
                    }, function (msgs, page/*, index*/) {
                        if (msgs === undefined) {
                            msgs = [];
                        }

                        msgs.push({
                            payload: page
                        });

                        return msgs;
                    });

                    const partsId = RED.util.generateId();
                    msgs = msgs.map(function (msg, index, msgs) {
                        ok = ok && msg.payload.ok;

                        const count = msgs.length;
                        msg.parts = {
                            id: partsId,
                            index,
                            count,
                            type: 'array'
                        };

                        if (index === count - 1) {
                            msg.complete = true;
                        }

                        return msg;
                    });
                } else {
                    msg.payload = await node.webClient.apiCall(methodName, options);

                    ok = msg.payload.ok;
                    msgs = [msg];
                }

                if (ok) {
                    node.status({fill: 'green', shape: 'dot', text: 'ok'});
                    _send([msgs, null]);
                } else {
                    node.status({fill: 'yellow', shape: 'dot', text: (msg.payload ? msg.payload.error : '(see logs)')});
                    _send([null, msgs]);
                }

                if (done) {
                    done();
                }
            } catch (e) {
                node.status({fill: 'red', shape: 'dot', text: e.code});

                if (done) {
                    done(e);
                } else {
                    node.error(e, msg);
                }
            }
        });
    }

    function ChannelLookupNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        const webClientConfig = RED.nodes.getNode(config.webClientConfig);
        this.webClient = webClientConfig.webClient;

        this.channel = config.channel;
        this.channelType = config.channelType;
        this.output = config.output;
        this.outputType = config.outputType;
        this.groupConversations = config.groupConversations;

        function conversationsByName() {
            return node.webClient.paginate('conversations.list', {types: 'public_channel,private_channel'}, function (page) {
                return false;
            }, function (channels, page) {
                if (channels === undefined) {
                    channels = {};
                }

                page.channels.forEach(function (channel) {
                    channels[channel.name] = channel.id;
                });

                return channels;
            });
        }

        function usersByName() {
            return node.webClient.paginate('users.list', {}, function (page) {
                return false;
            }, function (users, page) {
                if (users === undefined) {
                    users = {};
                }

                page.members.forEach(function (user) {
                    users[user.name] = user.id;
                });

                return users;
            });
        }

        function groupOfUserIds(names) {
            return names.length > 1 && names.filter(function (name) {
                return name[0] === 'U';
            }).length === names.length;
        }

        async function channelLikeIds(namesOrIds) {
            if (!Array.isArray(namesOrIds)) {
                namesOrIds = namesOrIds.split(/\s*,\s*/).filter(Boolean);
            }

            if (namesOrIds.length === 0) {
                return '';
            }

            const conversations = await conversationsByName();
            const users = await usersByName();

            namesOrIds = namesOrIds.map(function (name) {
                switch (name[0]) {
                    case '#':
                        return conversations[name.slice(1)];
                    case '@':
                        return users[name.slice(1)];
                    case 'C':
                    case 'G':
                    case 'D':
                    default:
                        return name;
                }
            }).filter(Boolean);

            if (node.groupConversations && groupOfUserIds(namesOrIds)) {
                const group = await node.webClient.conversations.open({users: namesOrIds.join(',')});
                if (group.ok) {
                    return group.channel.id;
                }
            }

            return namesOrIds.join(',');
        }

        this.on('input', async function (msg, _send, done) {
            let channel;
            if (node.channelType === 'nodeContext') {
                channel = node.context().get(node.channel);
            } else {
                channel = RED.util.evaluateNodeProperty(node.channel, node.channelType, node, msg);
            }

            if (channel) {
                // https://api.slack.com/methods/chat.postMessage#channels
                // Passing a "username" as a channel value is deprecated, along with the whole concept of usernames on Slack. Please always use channel-like IDs instead to make sure your message gets to where it's going.
                // https://api.slack.com/changelog/2017-09-the-one-about-usernames
                channel = await channelLikeIds(channel); // TODO cache lookup

                if (node.outputType === 'msg') {
                    RED.util.setMessageProperty(msg, node.output, channel, true);
                } else if (node.outputType === 'nodeContext') {
                    node.context().set(node.output, channel);
                } else if (node.outputType === 'flow') {
                    node.context().flow.set(node.output, channel);
                } else if (node.outputType === 'global') {
                    node.context().global.set(node.output, channel);
                } else {
                    msg.payload = channel;
                }
            }

            _send(msg);

            if (done) {
                done();
            }
        });
    }

    function EscapeTextNode(config) {
        RED.nodes.createNode(this, config);
        const node = this;

        this.text = config.text;
        this.textType = config.textType;
        this.output = config.output;
        this.outputType = config.outputType;

        this.on('input', async function (msg, _send, done) {
            let text;
            if (node.textType === 'nodeContext') {
                text = node.context().get(node.text);
            } else {
                text = RED.util.evaluateNodeProperty(node.text, node.textType, node, msg);
            }

            if (text) {
                // https://api.slack.com/reference/surfaces/formatting#escaping
                text = text.replace(/[&<>]/gm, function (match) {
                    switch (match) {
                        case '&':
                            return '&amp;';
                        case '<':
                            return '&lt;';
                        case '>':
                            return '&gt;';
                        default:
                            return match;
                    }
                });

                if (node.outputType === 'msg') {
                    RED.util.setMessageProperty(msg, node.output, text, true);
                } else if (node.outputType === 'nodeContext') {
                    node.context().set(node.output, text);
                } else if (node.outputType === 'flow') {
                    node.context().flow.set(node.output, text);
                } else if (node.outputType === 'global') {
                    node.context().global.set(node.output, text);
                } else {
                    msg.payload = text;
                }
            }

            _send(msg);

            if (done) {
                done();
            }
        });
    }

    RED.nodes.registerType('slack-webclient-config', WebClientConfig, {
        credentials: {
            token: {
                type: 'password'
            }
        }
    });
    RED.nodes.registerType('slack-webclient', WebClientNode);
    RED.nodes.registerType('slack-channel-lookup', ChannelLookupNode);
    RED.nodes.registerType('slack-escape-text', EscapeTextNode);
}
