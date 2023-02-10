import React from 'react';
import { app, authentication } from '@microsoft/teams-js';

interface JwtTokenHeader {
    typ: string
    alg: string
  }

class Auth extends React.Component<{}, { context: {}, token: string | null }> {

    constructor(props: any) {
        super(props);
        app.initialize();
        console.log('initialized');

        this.state = {
            context: {},
            token: null,
        }
    }

    async componentDidMount(): Promise<void> {
        const context: app.Context = await app.getContext();
        const token = await this.fetchTokens();

        this.setState({ context, token })
        
    }

    fetchTokens = async (): Promise<string> => {
        const token = await authentication.getAuthToken();
        console.log({token});
        // alert(JSON.stringify({token}));
        return token;
    }

    looksLikeJwtToken = (token: string | null): boolean => {
        if (!token) {
            return false
        }
        try {
          const header: JwtTokenHeader = JSON.parse(atob(token.split('.')[0]))
          return header.typ.toLowerCase() === 'jwt'
        } catch (e) {
          return false
        }
      }

    render(): React.ReactNode {
        const { context } = this.state;
        const contextArr = Object.keys(context).length ? Object.entries(context) : [];
        const contextContent = contextArr.map(([key, value], index) => {
            // alert(JSON.stringify({key, value}));
            return (
                <div key={`${index}-${key}`} style={{ marginBottom: '10px' }}>
                    <div>{key}:</div>
                    <div>{JSON.stringify(value)}</div>
                </div>
            )
        });
        // alert(JSON.stringify({ contextContent}));

        const flowType = this.looksLikeJwtToken(this.state.token) ? 'aad' : 'msa';

        return (
            <div>
                <h1>Auth</h1>

                <h2>Context</h2>
                <div>{contextContent}</div>

                <h2>Token</h2>
                <div>{this.state.token}</div>

                <h2>flowType</h2>
                <div>{flowType}</div>
            </div>
        );
    }
}

export { Auth };