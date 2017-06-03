# A Simple ImacrosWrapper written in Java 

## Requirements
Please ensure that you have the Jacob (Java COM Bridge) in your classpath while compiling this library. It can be downloaded from the link below.

[Download Jacob](http://danadler.com/jacob/)

## Example
```java
	ImacrosWrapper i = new ImacrosWrapper();
	i.iimSet("urlvar","http://google.com");
	Variant v = i.iimPlay("DemoImacros");
    	System.out.println(v);
	v = i.iimGetExtract(1);
    	System.out.println(v.getString());
	i.iimClose();
```
