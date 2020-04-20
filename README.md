# SimpleJsonCompare

Using Excel Json File Compare

- [Download](https://github.com/xiz2002/SimpleJsonCompare/releases/latest)

## Example

### 2 Files Compare
```
File_1.json
{"values":[{"a":1,"b":2,"c":3},{"d":"2000/01/01","e":[1,2,3,4,5,6],"f":"abcde"}]}

File_2.json
{"values":[{"a":1,"c":2,"b": 3}, {"d":0,"e":null,"f":[{"h":true}]}]}
```
↓

| Field_Name | File_1.Value | File_2.Value | isDiff(1 eq 2) |
| :--: | -- | -- | -- | 
| **values.1.a** | 1 | 1 | FALSE |
| **values.1.b** | 2 | 3 | TRUE |
| **values.1.c** | 3 | 2 | TRUE |
| **values.2.d** | 2000/01/01 | 0 | TRUE | 
| **values.2.e** | 1,2,3,4,5,6 | | TRUE | 
| **values.2.f** | abcde | | TRUE | 
| **values.2.f.1.h** | | TRUE | TRUE | 

### 3 Files Compare 
```
File_1.json
{"values":[{"a":1,"b":2,"c":3},{"d":"2000/01/01","e":[1,2,3,4,5,6],"f":"abcde"}]}

File_2.json
{"values":[{"a":1,"c":2,"b": 3}, {"d":0,"e":null,"f":[{"h":true}]}]}

File_3.json
[{"values":[1,2,3,4,5,6,7],"key":"key"}, true]
```
↓

| Field_Name | File_1.Value | File_2.Value | File_3.Value | isDiff(1 eq 3) | isDiff(2 eq 3) | isDiff(1 eq 2) |
| :--: | -- | -- | -- | -- | -- | -- |
| **values.1.a** | 1 | 1 |  | TRUE | TRUE | FALSE |
| **values.1.b** | 2 | 3 |  | TRUE | TRUE | TRUE |
| **values.1.c** | 3 | 2 |  | TRUE | TRUE | TRUE |
| **values.2.d** | 2000/01/01 | 0 |  | TRUE | TRUE | TRUE | 
| **values.2.e** | 1,2,3,4,5,6 | | | TRUE | FALSE | TRUE | 
| **values.2.f** | abcde | | | TRUE | FALSE | TRUE | 
| **values.2.f.1.h** | | TRUE |  | FALSE | TRUE | TRUE | 
| **1.values** | |  | 1,2,3,4,5,6,7 | TRUE | TRUE | FALSE | 
| **1.key** | |  | key | TRUE | TRUE | FALSE | 
| **2** |	 |  | TRUE | TRUE | TRUE | FALSE | 


## Dependency
  [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)

## License
  [MIT](LICENSE)
