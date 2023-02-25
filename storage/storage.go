package storage

type storage struct {
	driver     Driver
	bucket     string
	credential Credential
}


type Driver interface {
	// PutObject put object to storage
	PutObject(bucket, key string, data []byte) error
	// GetObject get object from storage
	GetObject(bucket, key string) ([]byte, error)
	// DeleteObject delete object from storage
	DeleteObject(bucket, key string) error
	// ListObjects list objects from storage
	ListObjects(bucket, prefix string) ([]string, error)
}

type Credential interface {
	// GetAccessKeyID get access key id
	GetAccessKeyID() string
	// GetSecretAccessKey get secret access key
	GetSecretAccessKey() string
}

// NewStorage create a new storage
func NewStorage(driver Driver, bucket string, credential Credential) *storage {
	return &storage{
		driver:     driver,
		bucket:     bucket,
		credential: credential,
	}
}

// PutObject put object to storage
func (s *storage) PutObject(key string, data []byte) error {
	return s.driver.PutObject(s.bucket, key, data)
}