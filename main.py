import pandas as pd
fuzzyAnggota =  {
  'servis': [],
  'harga': []
}

fuzzyAnggota['servis'].append({'buruk': [1, 40]})
fuzzyAnggota['servis'].append({'cukup': [41, 70]})
fuzzyAnggota['servis'].append({'baik': [71, 100]})

fuzzyAnggota['harga'].append({'murah': [1, 4]})
fuzzyAnggota['harga'].append({'cukup': [5, 7]})
fuzzyAnggota['harga'].append({'mahal': [8, 10]})

fuzzyRule = {
  ('buruk', 'mahal'): 'tidak rekomen',
  ('buruk', 'cukup'): 'tidak rekomen',
  ('buruk', 'murah'): 'rekomen',
  ('cukup', 'mahal'): 'tidak rekomen',
  ('cukup', 'cukup'): 'cukup rekomen',
  ('cukup', 'murah'): 'rekomen',
  ('baik', 'mahal'): 'rekomen',
  ('baik', 'cukup'): 'sangat rekomen',
  ('baik', 'murah'): 'sangat rekomen',
}

deffuzy = {'tidak rekomen': 0, "cukup rekomen": 50, "rekomen": 75, "sangat rekomen": 100}


def BuatFuzzy():
  data = fuzzyAnggota
  keanggotaan = {}
  for i in data:
      keanggotaan[i] = {}
      for j in data[i]:
          k = list(j.keys())[0]
          keanggotaan[i][k] = 0
  return keanggotaan

def ReadFromExcel():
  return pd.read_excel('./bengkel.xlsx')

def Fuzzification(fuzzy, x):
  anggota = BuatFuzzy()
  for j, data in enumerate(anggota[fuzzy].keys()):
      b, c = fuzzyAnggota[fuzzy][j][data]
      a, d = b - 1, c + 1
      if b <= x <= c:
          anggota[fuzzy][data] = 1
      elif a < x < b:
          anggota[fuzzy][data] = (x-a)/(b-a)
      elif c < x <= d:
          anggota[fuzzy][data] = (d-x)/(d-c)
  return anggota[fuzzy]

def FuzzificationData(excel):
  fuzzied = []
  for i in range(len(excel)):
      fuzzed = {}
      for j in fuzzyAnggota:
          fuzzed[j] = Fuzzification(j, excel[j][i])
      fuzzied.append(fuzzed)
  return fuzzied

def Inference(fuzzied):
  inferenced = []
  for fuzzed in fuzzied:
      result = {}
      keys = []
      for i in fuzzyRule:
          keys.append(i)
          result[fuzzyRule[i]] = 0
      for key in keys:
          output = fuzzyRule[key]
          minVal = fuzzed[list(fuzzed.keys())[0]][key[0]]
          for j, val in enumerate(fuzzed):
              minVal = min(minVal, fuzzed[val][key[j]])
          result[output] = max(minVal, result[output])
      inferenced.append(result)
  return inferenced

def Defuzzification(inferenced):
  results = []
  for inference in inferenced:
      numerator, denominator = 0, 0
      for output in deffuzy.keys():
          numerator += inference[output] * deffuzy[output]
          denominator += inference[output]
      results.append(numerator/denominator)
  
  data = ReadFromExcel()
  data['result'] = results
  return data

data = ReadFromExcel()
fuzzied = FuzzificationData(data)
inferenced = Inference(fuzzied)
defuzz = Defuzzification(inferenced)
print(defuzz.sort_values(by='result', ascending=False)[:10])

tmp = defuzz.sort_values(by='result', ascending=False)[:10]
tmp.to_excel("peringkat.xlsx", engine='openpyxl', index=False)