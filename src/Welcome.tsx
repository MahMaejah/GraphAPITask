import { Client, GraphRequestOptions, PageCollection, PageIterator } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import {
    Button,
    Container
  } from 'react-bootstrap';
  import { RouteComponentProps } from 'react-router-dom';
  import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
  import { useAppContext } from './AppContext';
  import { getMail } from './GraphService';
  import { User, Event } from 'microsoft-graph';
import { useEffect, useState } from 'react';

  export default function Welcome(props: RouteComponentProps) {
    const app = useAppContext();
    console.log(app)
    const [mails, setMails] = useState([])

    useEffect(() => {
      ( async () => {
        try {
          const mails = await getMail(app.authProvider)
          setMails(mails.value)
        } catch (error) {
          console.log(error)
        }
      })()
      
    }, [])

    let emailClicked = (p) => {
      console.log(p);
      p.target.style.display = "None";
      alert(p.target.nextSibling.innerText)
    }

    return (
      <div className="p-5 mb-4 bg-light rounded-3">
        <Container fluid>
          <AuthenticatedTemplate>
            <div>
              <h4>Welcome {app.user?.displayName || ''}!</h4>
              <p>Use the navigation bar at the top of the page to get started.</p>

              <div className="mail-list">
                  <h1>Your Mail</h1><hr />
                  <b><i>Unread Mail</i></b>
                  <div className="mail-list">
                    {mails.map((mail, index) => {
                      return (
                        <div key={index} onClick={(e) => emailClicked(e)} className="mail">
                          <h2 className="mail-subject">{index} {mail.subject}</h2>
                          <div className='mail-message'><p>{mail.bodyPreview}</p></div>
                        </div>
                      )
                    })}
                    
                  </div>
              </div>
            </div>
          </AuthenticatedTemplate>
          <UnauthenticatedTemplate>
            <Button color="primary" onClick={app.signIn!}>Click here to sign in</Button>
          </UnauthenticatedTemplate>
        </Container>
      </div>
    );
  }