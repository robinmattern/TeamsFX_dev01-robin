### Generative Artificial Intelligent Documents

1. **Create a remote GitHub repository**

	`- Goto URL: https://github.com/robinmattern/new`  
    `- Select: New repository`  
    `- - Enter Repository name: GAID_dev01-robin`  
    `- - Check: Add a README file`  
    `- - Add .gitignore template: Node`  
    `- - Choose a license: GNU General Public License v3.0`   
    `- - Click button: Create Repository`        

2. **Create a local Git Repository for VSCode**
	
	`# cd C:/Repos`  
	`# mkdir -p GAID_/dev01-robin`  
	`# cd GAID_/dev01-robin`  
	`# git init`  
    `# git remote add origin git@github-ram:robinmattern  GAID_dev01-robin.git`  
	`# echo '{ "folders": [ { "path": ".\\" } ], "settings": {} }' >GAID_dev01-robin.code-workspace`  
    `# code GAID_dev01-robin.code-workspace`  

3. **Publish your first commit**

    `- Create a new file: .gitignore`

        node_modules
        *_v[0-9]* 
        .env

    `- Commit changes with comment: c30620.01_Initial Setup`
    `- Publish Branch`


