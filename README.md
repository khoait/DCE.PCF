# 🚀 DynCRMExp Power Apps PCF Components Repository

![PowerApps PCF](https://img.shields.io/badge/PowerApps%20PCF%20Components-742774?logo=powerapps&logoColor=white)
[![GitHub all releases](https://img.shields.io/github/downloads/khoait/dce.pcf/total?style=social)](https://github.com/khoait/DCE.PCF/releases)
[![Sponsor on Github](https://img.shields.io/badge/Buy%20me%20a%20coffee-323330?logo=githubsponsors)](https://github.com/sponsors/khoait)

This Github repository contains a collection of React-based components that provide extensive functionalities with consistent UI/UX to Power Apps using the Fluent UI Design Language. The components are designed to be easy to use and highly customizable, allowing you to quickly build powerful and visually appealing Power Apps.

## 📚 Components Documentation

The components included in this repository are fully documented with examples and usage guidelines. You can find the documentation in the wiki page of the repository.

1. **[PolyLookup](https://github.com/khoait/DCE.PCF/wiki/PolyLookup)**: Multi-select lookup supporting different types of many-to-many relationships.
2. **[Lookdown](https://github.com/khoait/DCE.PCF/wiki/Lookdown)**: Lookup field as a dropdown with advanced features.

## ⚙️ Installation

To use these components in your Power Apps project, you'll need to download the `solution_managed.zip` file from the [Releases page](https://github.com/khoait/DCE.PCF/releases) and import to your environment. The solution contains all components in this repo.

Once you've installed the solution, you can select the components in the form designer.

## 🤝 Contribution

Thanks for showing your interest in the repo. All contributions are most welcome!

### Ways to contribute

- **Share**: Give it a star ⭐ and Share the repo with everyone who might be interested!
- **Give feedback**: Please share your experience and ideas.
- **Improve documentation**: Fix incomplete or missing docs, bad wording, examples or explanations.
- **Contribute to the codebase**: Propose new features or identify problems via GitHub Issues, or find an existing issue that you are interested in and work on it!
- **Sponsor my work**: If you find it helpful, please consider sponsoring the project to keep it going 💪

If you'd like to contribute to the code, please follow these steps:

1. Fork this repository
2. Create a new branch from a component branch you want to work on
3. Make your changes and commit them with descriptive commit messages
4. Submit a pull request

I'll review your changes and work with you to merge them into the main repository.

### Git branches

- **main**: This branch contains the latest features of the components. Please see [tags](https://github.com/khoait/DCE.PCF/tags) for version specific code

### Build and test individual component in dev

1. open the repo in VS Code
2. in Terminal, navigate to a component folder `cd PolyLookupComponent` for example
3. run `npm install` to install all required packages
4. run `npm run build` to build the component
5. to test the component, because the component needs to load metadata from Dataverse, we can't use the test harness by running `npm run start`. 
   
   Instead, we need to deploy the component to an environment by running 
   ```
   pac pcf push --publisher-prefix dce
   ```

   you'll need PowerApps CLI to run `pac` commands, please install it [here](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction)

### Build the component package for production

when you're ready, you can build a solution package that contains all components in the repo to deploy to a production environment.

1. in Terminal, navigate to folder `cd DCEPCF/solution`
2. for dotnet build, run  
   ```
   dotnet build --configuration Release
   ```

   for msbuild, run 
   
   ```
   msbuild /t:build /restore /p:configuration=Release
   ```
3. you can find an managed solution and an unmanaged solution will be created in the output folder under `solution/bin/Release`