# MSProjectEx - CSharp ![build](https://img.shields.io/badge/build-passing-blue)![lang](https://img.shields.io/badge/lang-.NET5-blue)![licence](https://img.shields.io/badge/licence-MIT-blue)
Sample about read from or write to MS Project 2016.



### Noted :notebook_with_decorative_cover:

> Here are two demo about how to read from or write to MS Project 2016

- _WIP_ - using `NetOffice` library and here is the origin repo [NetOffice](https://github.com/netoffice/NetOffice) ![status](https://img.shields.io/badge/process-WIP-blueviolet)
- _Done_ - using `MSProject` COM+ and `Office` library ![status](https://img.shields.io/badge/process-DONE-brightgreen)



### Getting Started :clap:

> Microsoft Visual Studio

- Create a .NET solution 
- Add a Console Application
- Add dependency
  - COM+ (Microsoft Office Project 16.0 Object Library)
    - Refer from this path `C:\Windows\assembly\GAC_MSIL\Policy.14.0.Microsoft.Office.Interop.MSProject\15.0.0.0__71e9bce111e9429c\Policy.14.0.Microsoft.Office.Interop.MSProject.dll` 
  - OFFICE (Program Libraries)
    - Refer from this path `C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL`
- Create a Class about Microsoft Project file informations (`TaskId` `TaskName` `Start` `End` `Duration` `ResourceName` ... )
- Create a MS Project Handle about common methods
- Now, you can perform a read from or write to MS Project 2016 file



### TODO

- NetOffice :construction:



### License :policeman:

<p>
<a href="LICENSE">MIT License</a>
</p>



### Sponsor

> If you think this repo is helpful to you and you happen to have extra money, then you might as well sponsor a cup of coffee for me, many thanks :pray:

Wechat or Alipay :coffee:

<p>
    <img height="100" src="https://gitee.com/jeremywuiot/img-res-all/raw/master/src/iie_shop/Wechat.png"/>
    <img height="100" src="https://gitee.com/jeremywuiot/img-res-all/raw/master/src/iie_shop/Alipay.png"/>
</p>

