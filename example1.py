# https://developer.atlassian.com/server/jira/platform/jira-rest-api-examples/#searching-for-issues-examples
# oauth exanple https://bitbucket.org/atlassianlabs/atlassian-oauth-examples/src/master/python/app.py

import base64
import oauth2 as oauth
import requests
import urllib.parse

from tlslite.utils import keyfactory


class SignatureMethod_RSA_SHA1(oauth.SignatureMethod):
    name = 'RSA-SHA1'
    url = f"https://oandacorp.atlassian.net/browse/W20-466?jql=project%20%3D%20W20%20AND%20status%20%3D%20Done%20AND%20assignee%20in%20(currentUser())%20order%20by%20created%20DESC"
    def signing_base(self, request, consumer, token):
        if not hasattr(request, 'normalized_url') or request.normalized_url is None:
            raise ValueError("Base URL for request is not set.")

        sig = (
            oauth.escape(request.method),
            oauth.escape(request.normalized_url),
            oauth.escape(request.get_normalized_parameters()),
        )

        key = '%s&' % oauth.escape(consumer.secret)
        if token:
            key += oauth.escape(token.secret)
        raw = '&'.join(sig)
        return key, raw

    def sign(self, request, consumer, token):
        """Builds the base signature string."""
        key, raw = self.signing_base(request, consumer, token)

        with open('../rsa.pem', 'r') as f:
            data = f.read()
        privateKeyString = data.strip()

        privatekey = keyfactory.parsePrivateKey(privateKeyString)
        signature = privatekey.hashAndSign(raw)

        return base64.b64encode(signature)


consumer_key = 'oauth-sample-consumer'
consumer_secret = 'dont_care'

request_token_url = 'http://localhost:8090/jira/plugins/servlet/oauth/request-token'
access_token_url = 'http://localhost:8090/jira/plugins/servlet/oauth/access-token'
authorize_url = 'http://localhost:8090/jira/plugins/servlet/oauth/authorize'

data_url = 'http://localhost:8090/jira/rest/api/2/issue/BULK-1'

consumer = oauth.Consumer(consumer_key, consumer_secret)
client = oauth.Client(consumer)

# Lets try to access a JIRA issue (BULK-1). We should get a 401.
resp, content = client.request(data_url, "GET")
if resp['status'] != '401':
    raise Exception("Should have no access!")

consumer = oauth.Consumer(consumer_key, consumer_secret)
client = oauth.Client(consumer)
client.set_signature_method(SignatureMethod_RSA_SHA1())

# Step 1: Get a request token. This is a temporary token that is used for
# having the user authorize an access token and to sign the request to obtain
# said access token.

resp, content = client.request(request_token_url, "POST")
if resp['status'] != '200':
    raise Exception("Invalid response %s: %s" % (resp['status'], content))

request_token = dict(urllib.parse.parse_qsl(content))

print("Request Token:    - oauth_token        = %s" % request_token['oauth_token'],
      "    - oauth_token_secret = %s" % request_token['oauth_token_secret'])


# Step 2: Redirect to the provider. Since this is a CLI script we do not
# redirect. In a web application you would redirect the user to the URL
# below.

print("Go to the following link in your browser:", "%s?oauth_token=%s" % (authorize_url, request_token['oauth_token']))


# After the user has granted access to you, the consumer, the provider will
# redirect you to whatever URL you have told them to redirect to. You can
# usually define this in the oauth_callback argument as well.
accepted = 'n'
while accepted.lower() == 'n':
    accepted = input('Have you authorized me? (y/n) ')
# oauth_verifier = input('What is the PIN? ')

# Step 3: Once the consumer has redirected the user back to the oauth_callback
# URL you can request the access token the user has approved. You use the
# request token to sign this request. After this is done you throw away the
# request token and use the access token returned. You should store this
# access token somewhere safe, like a database, for future use.
token = oauth.Token(request_token['oauth_token'],
                    request_token['oauth_token_secret'])
# token.set_verifier(oauth_verifier)
client = oauth.Client(consumer, token)
client.set_signature_method(SignatureMethod_RSA_SHA1())

resp, content = client.request(access_token_url, "POST")
access_token = dict(urllib.parse.parse_qsl(content))

print("Access Token:", "    - oauth_token        = %s" % access_token['oauth_token'],
      "    - oauth_token_secret = %s" % access_token['oauth_token_secret'],
      "You may now access protected resources using the access tokens above.")

# Now lets try to access the same issue again with the access token. We should get a 200!
accessToken = oauth.Token(access_token['oauth_token'], access_token['oauth_token_secret'])
client = oauth.Client(consumer, accessToken)
client.set_signature_method(SignatureMethod_RSA_SHA1())

resp, content = client.request(data_url, "GET")
if resp['status'] != '200':
    raise Exception("Should have access!")
