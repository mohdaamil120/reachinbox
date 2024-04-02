const { google } = require('googleapis');
const express = require('express');
const session = require('express-session');
const dotenv = require('dotenv');
const passport = require("passport")
const GoogleStrategy = require("passport-google-oauth").OAuth2Strategy
const axios = require('axios'); 
const { ClientSecretCredential } = require('@azure/identity');
// const { ClientSecretCredential, DeviceCodeCredential } = require('@azure/identity');
const { GraphClient } = require('@microsoft/microsoft-graph-client');
const  Queue  = require('bull');

dotenv.config();

const app = express();

app.use(session({
  secret: process.env.SESSION_SECRET_KEY,
  resave: false, 
  saveUninitialized: true
}));



const queue = new Queue('email-processing');


passport.use(new GoogleStrategy({
  clientID: process.env.GOOGLE_CLIENT_ID,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  callbackURL: "http://localhost:3000/auth/google/callback",
},
function(accessToken, refreshToken, profile, done) {
  // Set tokens in session
  const user = {
    profile: profile,
    tokens: {
      accessToken: accessToken,
      refreshToken: refreshToken
    }
  };
  return done(null, user);

}
));





// const googleOAuth2Client = new google.auth.OAuth2(
//   process.env.GOOGLE_CLIENT_ID,
//   process.env.GOOGLE_CLIENT_SECRET,
//   "http://localhost:3000/auth/google/callback"
// );

passport.serializeUser(function(user,done){
  done(null, user)
})


passport.deserializeUser(function(user,done){
  done(null, user)
})

app.use(passport.initialize())
app.use(passport.session())

app.get("/" ,(req,res) =>{
  res.send("Hello world")
})

app.get("/auth/google", 
  passport.authenticate("google", {scope : ["profile" ,"email"], accessType: 'offline'})

);



app.get("/auth/google/callback" ,
  passport.authenticate("google", {failureRedirect: "/login" }),
  function(req,res){

    const accessToken = req.user.tokens.accessToken;
    const refreshToken = req.user.tokens.refreshToken;

    req.session.googleTokens = {
      accessToken: accessToken,
      refreshToken: refreshToken
    };

    res.redirect("/emails");
    
  }
)



const outlookCredential = new ClientSecretCredential(
  
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_CLIENT_SECRET
);


const openAIKey = process.env.OPENAI_API_KEY;


app.get('/auth/outlook', async (req, res) => {
  try {
    const authorizationUrl = await outlookCredential.getAuthorizationCodeUrl({
      redirectUri: 'http://localhost:3000/auth/outlook/callback',
      scopes: ['openid', 'profile', 'offline_access', 'https://outlook.office.com/mail.read']
    });
    res.redirect(authorizationUrl);
  } catch (error) {
    console.error('Error initiating Outlook authentication:', error.message);
    res.status(500).send('Error initiating Outlook authentication.');
  }
});

app.get('/auth/outlook/callback', async (req, res) => {
  const { code } = req.query;

  try {
    const tokenResponse = await outlookCredential.getToken('https://outlook.office.com/mail.read', code);
    const accessToken = tokenResponse.token.accessToken;
    req.session.outlookAccessToken = accessToken;
    res.send('Outlook Authentication successful. You can now access Outlook.');
  } catch (error) {
    console.error('Error authenticating with Outlook:', error.message);
    res.status(500).send('Error authenticating with Outlook.');
  }
});



app.get('/emails', async (req, res) => {
  try {
    const gmailAuth = new google.auth.OAuth2({
      clientId: "451279009887-p4i8n6s0ead6uspl06cl45h2cmb1oon6.apps.googleusercontent.com",
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      redirectUri:"http://localhost:3000/auth/google/callback",
    });

      if (req.session.googleTokens) {
        gmailAuth.setCredentials(req.session.googleTokens);
      } else {
        return res.redirect('/auth/google');
      }
    console.log('gmailAuth credentials:', gmailAuth.credentials);
    const gmail = google.gmail({ version: 'v1', auth: gmailAuth });
    const googleResponse = await gmail.users.messages.list({ userId: 'me', q:"is:unread" });
    const googleEmails = googleResponse.data;

    const outlookClient = GraphClient.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return req.session.outlookAccessToken;
        },
      },
    });
    const outlookResponse = await outlookClient.api('/me/messages').get();
    const outlookEmails = outlookResponse.value;

    const emails = [...googleEmails, ...outlookEmails];

    const emailContents = emails.map(email => email.snippet);

    const openAIResponse = await axios.post(
      'https://api.openai.com/v1/engines/davinci-1/completions',
      {
        prompt: emailContents.join('\n\n'),
        max_tokens: 100
      },
      {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${openAIKey}`
        }
      }
    );

    const analyzedContext = openAIResponse.data.choices.map(choice => choice.text.trim());

    res.json({ emails: emails, analyzedContext: analyzedContext });
  
    await queue.add('process-emails', { emails, analyzedContext });
    console.log("googleEmails : "+ googleEmails)
    res.send('Emails added to the processing queue.');

  } catch (error) {
    console.error('Error fetching and analyzing emails:', error.message);
    res.status(500).send('Error fetching and analyzing emails.');
  }
});

queue.process('process-emails', async (job) => {
  const { emails, analyzedContext } = job.data;

  try {
    const categorizedEmails = categorizeEmails(emails, analyzedContext);

    await sendAutomatedReplies(categorizedEmails);

    console.log('Emails processed and automated replies sent successfully.');
    return 'Emails processed and automated replies sent successfully.';
  } catch (error) {
    console.error('Error processing emails and sending automated replies:', error.message);
    throw new Error('Error processing emails and sending automated replies.');
  }
});

function categorizeEmails(emails, analyzedContext) {
  const categorizedEmails = emails.map(email => {
    if (analyzedContext.includes('interested')) {
      return { email: email, label: 'Interested' };
    } else if (analyzedContext.includes('not interested')) {
      return { email: email, label: 'Not Interested' };
    } else {
      return { email: email, label: 'More Information' };
    }
  });

  return categorizedEmails;
}

async function sendAutomatedReplies(categorizedEmails) {
  for (const email of categorizedEmails) {
    let automatedReply = '';
    if (email.label === 'Interested') {
      automatedReply = "Thank you for your interest! Would you be available for a demo call? Let us know your preferred time.";
    } else if (email.label === 'Not Interested') {
      automatedReply = "Thank you for considering us. If you change your mind, feel free to reach out!";
    } else {
      automatedReply = "We appreciate your inquiry. Can you provide more details so we can assist you better?";
    }
    
    await sendEmail(email.email.sender, email.email.recipient, email.email.subject, automatedReply);
  }
}

async function sendEmail(sender, recipient, subject, body) {
  try {
    const gmailAuth = new google.auth.OAuth2({
      clientId: "451279009887-p4i8n6s0ead6uspl06cl45h2cmb1oon6.apps.googleusercontent.com",
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      redirectUri: "http://localhost:3000/auth/google/callback",
      credentials: req.session.googleTokens
    });
    const gmail = google.gmail({ version: 'v1', auth: gmailAuth });
    const message = `
      From: ${sender}\r\n
      To: ${recipient}\r\n
      Subject: ${subject}\r\n\r\n
      ${body}
    `;
    const encodedMessage = Buffer.from(message).toString('base64');
    const emailToSend = await gmail.users.messages.send({
      userId: 'me',
      requestBody: {
        raw: encodedMessage
      }
    });

    console.log('Automated email sent successfully:', emailToSend.data);
    return emailToSend.data;
  } catch (error) {
    console.error('Error sending automated email:', error.message);
    throw new Error('Error sending automated email.');
  }
}




app.post('/automated-replies', async (req, res) => {
  try {
    const automatedReply = "Thank you for your email. Our team will get back to you shortly.";

    const gmailAuth = new google.auth.OAuth2({
      clientId: "451279009887-p4i8n6s0ead6uspl06cl45h2cmb1oon6.apps.googleusercontent.com",
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      redirectUri: "http://localhost:3000/auth/google/callback",
      credentials: req.session.googleTokens
    });
    const gmail = google.gmail({ version: 'v1', auth: gmailAuth });
    const message = `From: ${req.body.sender}\r\nTo: ${req.body.recipient}\r\nSubject: ${req.body.subject}\r\n\r\n${automatedReply}`;
    const encodedMessage = Buffer.from(message).toString('base64');
    const emailToSend = await gmail.users.messages.send({
      userId: 'me',
      requestBody: {
        raw: encodedMessage
      }
    });

    res.json({ reply: automatedReply, messageId: emailToSend.data.id });
  } catch (error) {
    console.error('Error generating automated reply:', error.message);
    res.status(500).send('Error generating automated reply.');
  }
});



const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
