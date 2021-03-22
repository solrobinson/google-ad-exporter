const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');

const SCOPES = [
    'https://www.googleapis.com/auth/admin.directory.orgunit',
    'https://www.googleapis.com/auth/admin.directory.orgunit.readonly',
];
const TOKEN_PATH = './authCreds/token.json';

fs.readFile('./authCreds/credentials.json', async (err, content) => {
    if (err) return console.error('Error loading client secret file', err);

    console.log("Authorizing Google Workspace Token....");
    const authResult = await authorize(JSON.parse(content));
    if(!authResult.newToken) {
        console.log("Authorized Token, collecting organisational units...");
        const orgGetRes = await listOUs(authResult.client);
        if(orgGetRes.success) {
            await handleOUs(orgGetRes.array[0]);
        } else {
            console.error(orgGetRes.msg);
        }
    }
    else {
        console.log("Token could not be authorised; Token missing, creating new token....");
        const result = await getNewToken(authResult.client);
        if(result.success) {
            let orgGetRes = await listOUs(authResult.client);
            if(orgGetRes.success) {
                await handleOUs(orgGetRes.array[0]);
            } else {
                console.error(orgGetRes.msg);
            }
        }
            
    }
});

const handleOUs = (array) => {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
    rl.question('Please specify the root DN: ', async dn => {
        rl.close();
        if(dn.length <= 3) {
            return console.error('Invalid Distribution Name. (ex. CN=domain,CN=com)');
        }
        let psData = [];

        fs.writeFile('createOrgUnits.ps1', '', (err) => {
            if(err) throw err;
        });

        let created = []

        await Promise.all(array.map(ou => {
            let path = ou.orgUnitPath.split('/')
            path = path.splice(1, path.length)
            return path.map(ouName => {
                let genRes = generatePSLine(ouName, array, dn)
                genRes.then(res => {
                    let foundCreated = false
                    created.forEach(names => {
                        if(names == res.name && foundCreated != true)
                            foundCreated = true;
                    });
                    if(foundCreated == false) {
                        created.push(res.name);
                        fs.appendFile('createOrgUnits.ps1', res.line + "\n", (err) => {
                            if(err) throw err;
                        })
                    }
                })
            })
        }))
        fs.readFile('createOrgUnits.ps1', (err, data) => {
            if(err) throw err;
            if(data.length < 2) {
                console.error("There was an issue when generating the powershell script!\nThis could be due to no/unavailable organisation data from Google Workspace.");
            } else {
                console.log("Powershell Script Generated!");
                fs.appendFile('createOrgUnits.ps1', 'Read-Host -Prompt "Press Enter to exit"', err => {
                    if(err) throw err;
                })
            }
        })
    })
}

const generatePSLine = (ouName, ouArray, dn) => {
    return new Promise(async (res, rej) => {
        let ouInfo = getOUInformation(ouName, ouArray);
        let path = ouInfo.orgUnitPath.split('/')
        path = path.splice(1, path.length)
        path.reverse()
        let formattedPath = ""
        path.forEach(pathName => {
            if(pathName != ouName)
                formattedPath = formattedPath + 'OU=' + pathName + ','
        });
        let psLine = 'New-ADOrganizationalUnit -Name "' + ouInfo.name + '" -Path "' + formattedPath + dn + '"'
        res({
            line: psLine,
            name: ouName
        })
    })
}

const getOUInformation = (ouName, ouArray) => {
    let info = []
    ouArray.forEach(ou => {
        if(ou.name == ouName) {
            info = ou;
        }
    });
    return info
}

const authorize = (credentials) => {
    return new Promise((res, rej) => {
        const {client_secret, client_id, redirect_uris} = credentials.installed;
        const oauth2Client = new google.auth.OAuth2(
            client_id, client_secret, redirect_uris[0]);

        fs.readFile(TOKEN_PATH, (err, token) => {
            if(err) {
                return res({
                    newToken: true,
                    client: oauth2Client
                })
            }
            oauth2Client.credentials = JSON.parse(token);
            return res({
                newToken: false,
                client: oauth2Client
            });
        });
    })
}

const getNewToken = (oauth2Client) => {
    return new Promise((res, rej) => {
        const authUrl = oauth2Client.generateAuthUrl({
            access_type: 'offline',
            scope: SCOPES,
        });
        console.log('Authorize this app by visiting this url:', authUrl);
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout,
        });
        rl.question('Enter the code from that page here: ', (code) => {
            rl.close();
            oauth2Client.getToken(code, async (err, token) => {
                if(err) {
                    return rej({
                        error: true,
                        msg: "An error occured during authorization: " + err.message
                    });
                }
                oauth2Client.credentials = token;
                await storeToken(token);
                return res({success: true});
            });
        });
    })
}

const storeToken = token => {
    return new Promise((res, rej) => {
        fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
            if(err) {
                return rej({
                    error: true,
                    msg: `Token not stored to ${TOKEN_PATH}` + err
                })
            }
            return res(console.log(`Token stored to ${TOKEN_PATH}`));
        });
    })
}

const listOUs = auth => {
    return new Promise((resolve, reject) => {
        const service = google.admin({version: 'directory_v1', auth});
        let nameArray = [];
        service.orgunits.list({
            customerId: 'my_customer',
            orgUnitPath: '/',
            type: 'all'
        }, (err, res) => {
            if(err) {
                return reject({
                    error: true,
                    msg: 'The API returned an error:' + err.message
                })
            }
            nameArray.push(res.data.organizationUnits);
            resolve({success: true, array: nameArray})
        })
    })
}