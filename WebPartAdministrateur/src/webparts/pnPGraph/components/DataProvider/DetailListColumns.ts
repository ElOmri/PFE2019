import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
export const _columns: IColumn[] = [
    {
      key: 'photo',
      name: 'Photo de profil',
      fieldName: 'photo',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
  
    }, {
      key: 'name',
      name: 'Nom',
      fieldName: 'displayName',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    }, {
      key: 'jobTitle',
      name: 'Fonction',
      fieldName: 'jobTitle',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'officeLocation',
      name: 'Departement',
      fieldName: 'officeLocation',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'userPrincipalName',
      name: 'adresse mail',
      fieldName: 'userPrincipalName',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
  
    }
  ];