import cognitive_face as CF
from global_variables import personGroupId, Key, endpoint


CF.BaseUrl.set(endpoint)
CF.Key.set(Key)


res = CF.person_group.train(personGroupId)
print(res)
